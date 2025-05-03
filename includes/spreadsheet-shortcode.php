<?php

add_shortcode('spreadsheet', function ($atts) {
    $shortcode_start_time = microtime(true);

    $atts = shortcode_atts([
        'id' => '',
        'max-rows' => 0,
    ], $atts);

    $table_id = sanitize_text_field($atts['id']);
    $max_rows = intval($atts['max-rows']);

    if (empty($table_id)) {
        return '<p><strong>Error:</strong> Missing table ID.</p>';
    }

    // Render the spreadsheet table using utility
    $output = esv_render_spreadsheet_table($table_id, [
        'max_rows' => $max_rows,
        'pagination' => true,
        'sticky_header' => true,
        'add_context_menu' => true,
    ]);

    $shortcode_end_time = microtime(true);

    return($output);
});
