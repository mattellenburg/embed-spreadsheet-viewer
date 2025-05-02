<?php
/**
 * Excel Values Extractor Function
 * 
 * Extracts calculated values from an Excel XLSX file and creates a new XLSX
 * with only the values (no formulas). This mimics "Paste Special > Values" in Excel.
 */
function extract_excel_values($excel_url, $sheet_name, $prefix = '', $post_id = null) {
    if (empty($excel_url) || empty($sheet_name)) {
        return new WP_Error('invalid_input', 'Excel URL and sheet name are required.');
    }

    $wp_upload_dir = wp_upload_dir();
    $upload_dir = trailingslashit($wp_upload_dir['basedir']) . 'excel-extractor/';

    if (!file_exists($upload_dir)) {
        wp_mkdir_p($upload_dir);
        file_put_contents($upload_dir . '.htaccess', 'deny from all');
    }

    try {
        $timestamp = time();
        $original_base = pathinfo(parse_url($excel_url, PHP_URL_PATH), PATHINFO_FILENAME);
        $original_file = $upload_dir . "{$original_base}.xlsx";
        $values_file   = $upload_dir . "{$original_base}_values_only.xlsx";

        if ($post_id) {
            $old_original = get_post_meta($post_id, 'esv_original_path', true);
            $old_flattened = get_post_meta($post_id, 'esv_flattened_path', true);
        
            if ($old_original && file_exists($old_original)) unlink($old_original);
            if ($old_flattened && file_exists($old_flattened)) unlink($old_flattened);
        }
        
        $download_result = download_excel_file($excel_url, $original_file);
        if (is_wp_error($download_result)) {
            return $download_result;
        }

        $extractor = new Excel_Values_Extractor($original_file, $sheet_name, $values_file);
        $result = $extractor->extract();

        if (!$result) {
            return new WP_Error('extraction_failed', 'Failed to extract values from Excel file.');
        }

        $flattened_url = str_replace($wp_upload_dir['basedir'], $wp_upload_dir['baseurl'], $values_file);

        if ($post_id) {
            update_post_meta($post_id, 'esv_original_path', $original_file);
            update_post_meta($post_id, 'esv_flattened_path', $values_file);
            update_post_meta($post_id, 'esv_flattened_url', $flattened_url);
        }        

        return [
            'original_file' => basename($original_file),
            'values_file'   => basename($values_file),
            'download_url'  => $flattened_url,
            'timestamp'     => $timestamp,
            'file_path'     => $values_file
        ];
    } catch (Exception $e) {
        return new WP_Error('extraction_error', $e->getMessage());
    }
}

function download_excel_file($url, $destination) {
    $context = stream_context_create(['http' => ['timeout' => 300]]);
    $response = wp_remote_get($url, [
        'timeout'  => 300,
        'stream'   => true,
        'filename' => $destination
    ]);

    if (is_wp_error($response)) return $response;
    $code = wp_remote_retrieve_response_code($response);
    if ($code !== 200) return new WP_Error('download_failed', "HTTP error: $code");

    return true;
}

class Excel_Values_Extractor {
    private $inputFile, $sheetName, $outputFile, $tempDir;
    private $ns = [
        'main' => 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
        'r'    => 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    ];

    public function __construct($inputFile, $sheetName, $outputFile) {
        $this->inputFile = $inputFile;
        $this->sheetName = $sheetName;
        $this->outputFile = $outputFile;
        $this->tempDir = sys_get_temp_dir() . '/excel_extractor_' . uniqid();
    }

    public function extract() {
        try {
            if (!file_exists($this->tempDir)) mkdir($this->tempDir, 0777, true);
            $this->extractZip();
            $sheetId = $this->getSheetId();
            if (!$sheetId) throw new Exception("Sheet '$this->sheetName' not found");
            $this->processSheet($sheetId);
            $this->createNewXlsx();
            return true;
        } catch (Exception $e) {
            error_log("Extractor error: " . $e->getMessage());
            return false;
        } finally {
            $this->cleanupTempFiles();
        }
    }

    private function extractZip() {
        $zip = new ZipArchive();
        if ($zip->open($this->inputFile) !== true) {
            throw new Exception("Failed to open XLSX file");
        }
        $zip->extractTo($this->tempDir);
        $zip->close();
    }

    private function getSheetId() {
        $xml = simplexml_load_file($this->tempDir . '/xl/workbook.xml');
        $xml->registerXPathNamespace('main', $this->ns['main']);
        foreach ($xml->xpath('//main:sheets/main:sheet') as $sheet) {
            if ((string)$sheet['name'] === $this->sheetName) {
                return (string)$sheet['sheetId'];
            }
        }
        return null;
    }

    private function processSheet($sheetId) {
        $relsXml = file_get_contents($this->tempDir . '/xl/_rels/workbook.xml.rels');
        $rels = new SimpleXMLElement($relsXml);
        $sheetPath = null;
    
        foreach ($rels->Relationship as $rel) {
            if (strpos((string)$rel['Target'], 'sheet' . $sheetId . '.xml') !== false) {
                $sheetPath = 'xl/' . (string)$rel['Target'];
                break;
            }
        }
    
        if (!$sheetPath) throw new Exception("Sheet file for ID $sheetId not found");
    
        $sheetXmlPath = $this->tempDir . '/' . $sheetPath;
        $sheetXml = file_get_contents($sheetXmlPath);
        $dom = new DOMDocument();
        $dom->preserveWhiteSpace = false;
        $dom->formatOutput = false;
        $dom->loadXML($sheetXml);
        $xpath = new DOMXPath($dom);
        $xpath->registerNamespace('main', $this->ns['main']);
    
        $formulaCells = $xpath->query('//main:c[main:f]');
    
        foreach ($formulaCells as $cellNode) {
            /** @var DOMElement $cellNode */
            if (!($cellNode instanceof DOMElement)) continue;
    
            // ✅ Only remove <f> (formula) tags.
            foreach ($xpath->query('./main:f', $cellNode) as $f) {
                $cellNode->removeChild($f);
            }
        }
    
        $dom->save($sheetXmlPath);
    }    

    private function createNewXlsx() {
        $zip = new ZipArchive();
        if ($zip->open($this->outputFile, ZipArchive::CREATE | ZipArchive::OVERWRITE) !== true) {
            throw new Exception("Cannot write new XLSX file");
        }
        $this->addDirToZip($zip, $this->tempDir, '');
        $zip->close();
    }

    private function addDirToZip($zip, $dir, $zipPath) {
        $handle = opendir($dir);
        while (false !== ($entry = readdir($handle))) {
            if ($entry == '.' || $entry == '..') continue;
            $path = $dir . '/' . $entry;
            $zipEntryPath = ($zipPath === '') ? $entry : $zipPath . '/' . $entry;
            if (is_dir($path)) {
                $zip->addEmptyDir($zipEntryPath);
                $this->addDirToZip($zip, $path, $zipEntryPath);
            } else {
                $zip->addFile($path, $zipEntryPath);
            }
        }
        closedir($handle);
    }

    private function cleanupTempFiles() {
        $this->deleteDir($this->tempDir);
    }

    private function deleteDir($dir) {
        if (!file_exists($dir)) return;
        foreach (array_diff(scandir($dir), ['.', '..']) as $file) {
            $path = $dir . '/' . $file;
            is_dir($path) ? $this->deleteDir($path) : unlink($path);
        }
        rmdir($dir);
    }
}

// Auto-hooked after spreadsheet is saved
add_action('esv_after_spreadsheet_save', function ($url, $sheet, $post_id) {
    if (empty($url) || empty($sheet)) return;
    error_log("📦 Flatten triggered for: $url (sheet: $sheet)");

    $result = extract_excel_values($url, $sheet, 'flattened_', $post_id);

    if (is_wp_error($result)) {
        error_log('❌ Flatten error: ' . $result->get_error_message());
    } else {
        error_log('✅ Flattened file created: ' . $result['values_file']);
    }
}, 10, 3);

add_action('before_delete_post', function ($post_id) {
    if (get_post_type($post_id) !== 'esv_spreadsheet') return;

    $flattened_path = get_post_meta($post_id, 'esv_flattened_path', true);
    if ($flattened_path && file_exists($flattened_path)) {
        unlink($flattened_path);
    }
}, 10, 1);
