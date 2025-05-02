<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

add_action('wp_ajax_esv_preview_spreadsheet', function () {
    check_ajax_referer('esv_ajax_preview', 'security');

    $mode = sanitize_text_field($_POST['mode'] ?? 'worksheet');
    $table_id = sanitize_text_field($_POST['table_id'] ?? '');
    $worksheet_name = sanitize_text_field($_POST['worksheet'] ?? 'Sheet1');

    if (empty($table_id)) {
        wp_send_json_error('Table ID is missing.');
    }

    // Get post by table_id
    $post = get_posts([
        'post_type' => 'esv_spreadsheet',
        'meta_key' => 'esv_table_id',
        'meta_value' => $table_id,
        'posts_per_page' => 1,
    ]);

    if (empty($post)) {
        wp_send_json_error("Spreadsheet with ID '$table_id' not found.");
    }

    $post = $post[0];
    $flattened_path = get_post_meta($post->ID, 'esv_flattened_path', true);

    if (empty($flattened_path) || !file_exists($flattened_path)) {
        wp_send_json_error("Flattened spreadsheet file not found for '$table_id'. Try saving the spreadsheet again.");
    }

    try {
        $reader = IOFactory::createReaderForFile($flattened_path);
        $reader->setReadDataOnly(true);
        $reader->setLoadSheetsOnly([$worksheet_name]);
        $spreadsheet = $reader->load($flattened_path);
    } catch (Exception $e) {
        error_log('Spreadsheet load error: ' . $e->getMessage());
        wp_send_json_error('Failed to load flattened spreadsheet: ' . esc_html($e->getMessage()));
    }

    $sheet = $spreadsheet->getSheetByName($worksheet_name);
    if (!$sheet) {
        wp_send_json_error("Worksheet '$worksheet_name' not found.");
    }

    ob_start();

    // --- Branch behavior based on mode ---
    if ($mode === 'worksheet') {
        echo '<h3>Worksheet Preview (Raw)</h3>';
        echo '<div class="esv-preview-scroll">';
        echo '<table class="wp-list-table widefat fixed striped">';
        echo '<thead><tr><th>Row #</th>';

        $highestColumnIndex = Coordinate::columnIndexFromString($sheet->getHighestColumn());
        $highestRow = $sheet->getHighestRow();

        for ($col = 1; $col <= $highestColumnIndex; $col++) {
            echo '<th>' . esc_html(Coordinate::stringFromColumnIndex($col)) . '</th>';
        }
        echo '</tr></thead><tbody>';

        $row_limit = min($highestRow, 10);
        for ($row = 1; $row <= $row_limit; $row++) {
            echo '<tr><td>' . $row . '</td>';
            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $cell = $sheet->getCell(Coordinate::stringFromColumnIndex($col) . $row);
                echo '<td>' . esc_html($cell->getFormattedValue()) . '</td>';
            }
            echo '</tr>';
        }
        echo '</tbody></table>';
        echo '</div>'; // Close scroll container

    } elseif ($mode === 'table') {
        echo '<h3>Table Preview (Formatted)</h3>';

        echo esv_render_spreadsheet_table($table_id, [
            'max_rows' => 10,
            'pagination' => false,
            'context_menu' => false,
            'sticky_header' => false,
        ]);
    } else {
        wp_send_json_error('Invalid preview mode.');
    }

    $output = ob_get_clean();
    wp_send_json_success($output);
});

add_action('wp_ajax_esv_retry_flatten', function () {
    check_ajax_referer('esv_ajax_preview', 'nonce');

    $url = sanitize_text_field($_POST['url'] ?? '');
    $worksheet = sanitize_text_field($_POST['worksheet'] ?? 'Sheet1');
    $post_id = intval($_POST['post_id'] ?? 0);

    if (!$url || !$worksheet || !$post_id) {
        wp_send_json_error(['message' => 'Missing required data.']);
    }

    $result = extract_excel_values($url, $worksheet, 'flattened_', $post_id);

    if (is_wp_error($result)) {
        wp_send_json_error(['message' => $result->get_error_message()]);
    }

    wp_send_json_success(['path' => $result['file_path']]);
});

// For non-logged-in users
add_action('wp_ajax_nopriv_esv_retry_flatten', 'esv_retry_flatten_callback');

// Define the shared callback
function esv_retry_flatten_callback() {
    check_ajax_referer('esv_ajax_preview', 'nonce');

    $url      = esc_url_raw($_POST['url'] ?? '');
    $sheet    = sanitize_text_field($_POST['sheet'] ?? 'Sheet1');
    $post_id  = intval($_POST['post_id']);

    if (!$url || !$post_id) {
        wp_send_json_error(['message' => 'Missing URL or post ID.']);
    }

    $result = extract_excel_values($url, $sheet, 'flattened_', $post_id);

    if (is_wp_error($result)) {
        wp_send_json_error(['message' => $result->get_error_message()]);
    } else {
        wp_send_json_success(['message' => 'Flattened spreadsheet updated.']);
    }
}