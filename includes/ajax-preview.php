<?php
if ( ! defined( 'ABSPATH' ) ) exit;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

add_action('wp_ajax_esv_preview_spreadsheet', function () {
    check_ajax_referer('esv_ajax_preview', 'security');

    // Properly validate, unslash, and sanitize mode
    $mode = 'worksheet';
    if (isset($_POST['mode'])) {
        $mode = sanitize_text_field(wp_unslash($_POST['mode']));
    }

    // Properly validate, unslash, and sanitize table_id
    $table_id = '';
    if (isset($_POST['table_id'])) {
        $table_id = sanitize_text_field(wp_unslash($_POST['table_id']));
    }

    // Properly validate, unslash, and sanitize worksheet
    $worksheet_name = 'Sheet1';
    if (isset($_POST['worksheet'])) {
        $worksheet_name = sanitize_text_field(wp_unslash($_POST['worksheet']));
    }

    if (empty($table_id)) {
        wp_send_json_error('Table ID is missing.');
    }

    // Get post by table_id
    // phpcs:ignore WordPress.DB.SlowDBQuery.slow_db_query_meta_key, WordPress.DB.SlowDBQuery.slow_db_query_meta_value
    $post = get_posts([
        'post_type' => 'esv_spreadsheet',
        'posts_per_page' => 1,
        'meta_query'     => [
            [
                'key'     => 'esv_table_id',
                'value'   => $table_id,
                'compare' => '='
            ]
        ]
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
            echo '<tr><td>' . esc_attr($row) . '</td>';
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

        echo wp_kses_post(esv_render_spreadsheet_table($table_id, [
            'max_rows' => 10,
            'pagination' => false,
            'context_menu' => false,
            'sticky_header' => false,
        ]));
    } else {
        wp_send_json_error('Invalid preview mode.');
    }

    $output = ob_get_clean();
    wp_send_json_success($output);
});

add_action('wp_ajax_esv_retry_flatten', function () {
    check_ajax_referer('esv_ajax_preview', 'nonce');

    // Properly validate, unslash, and sanitize url
    $url = '';
    if (isset($_POST['url'])) {
        $url = sanitize_text_field(wp_unslash($_POST['url']));
    }
    
    // Properly validate, unslash, and sanitize worksheet
    $worksheet = 'Sheet1';
    if (isset($_POST['worksheet'])) {
        $worksheet = sanitize_text_field(wp_unslash($_POST['worksheet']));
    }
    
    // Properly validate post_id
    $post_id = 0;
    if (isset($_POST['post_id'])) {
        $post_id = intval($_POST['post_id']);
    }

    if (!$url || !$worksheet || !$post_id) {
        wp_send_json_error(['message' => 'Missing required data.']);
    }

    $result = esv_extract_excel_values($url, $worksheet, 'flattened_', $post_id);

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

    // Properly validate, unslash, and sanitize url
    $url = '';
    if (isset($_POST['url'])) {
        $url = esc_url_raw(wp_unslash($_POST['url']));
    }
    
    // Properly validate, unslash, and sanitize sheet
    $sheet = 'Sheet1';
    if (isset($_POST['sheet'])) {
        $sheet = sanitize_text_field(wp_unslash($_POST['sheet']));
    }
    
    // Properly validate post_id
    $post_id = 0;
    if (isset($_POST['post_id'])) {
        $post_id = intval($_POST['post_id']);
    }

    if (!$url || !$post_id) {
        wp_send_json_error(['message' => 'Missing URL or post ID.']);
    }

    $result = esv_extract_excel_values($url, $sheet, 'flattened_', $post_id);

    if (is_wp_error($result)) {
        wp_send_json_error(['message' => $result->get_error_message()]);
    } else {
        wp_send_json_success(['message' => 'Flattened spreadsheet updated.']);
    }
}