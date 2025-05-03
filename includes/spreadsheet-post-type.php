<?php

function esv_register_spreadsheet_cpt() {
    register_post_type('esv_spreadsheet', [
        'labels' => [
            'name'               => __('Spreadsheets', 'embed-spreadsheet-viewer'),
            'singular_name'      => __('Spreadsheet', 'embed-spreadsheet-viewer'),
            'add_new'            => __('Add New Spreadsheet', 'embed-spreadsheet-viewer'),
            'add_new_item'       => __('Add New Spreadsheet', 'embed-spreadsheet-viewer'),
            'edit_item'          => __('Edit Spreadsheet', 'embed-spreadsheet-viewer'),
            'new_item'           => __('New Spreadsheet', 'embed-spreadsheet-viewer'),
            'view_item'          => __('View Spreadsheet', 'embed-spreadsheet-viewer'),
            'search_items'       => __('Search Spreadsheets', 'embed-spreadsheet-viewer'),
            'not_found'          => __('No spreadsheets found', 'embed-spreadsheet-viewer'),
            'not_found_in_trash' => __('No spreadsheets found in Trash', 'embed-spreadsheet-viewer'),
            'menu_name'          => __('Spreadsheets', 'embed-spreadsheet-viewer'),
        ],
        'public' => false,
        'show_ui' => true,
        'menu_icon' => 'dashicons-media-spreadsheet',
        'supports' => ['title'],
        'show_in_menu' => 'esv_spreadsheet_menu'
    ]);
}
add_action('init', 'esv_register_spreadsheet_cpt');

// Add custom columns
function esv_spreadsheet_custom_columns($columns) {
    unset($columns['date']);
    $columns['esv_table_id'] = __('Table ID', 'embed-spreadsheet-viewer');
    $columns['esv_url'] = __('Spreadsheet URL', 'embed-spreadsheet-viewer');
    $columns['esv_flattened'] = __('Last Update Date', 'embed-spreadsheet-viewer');
    $columns['date'] = __('Date', 'embed-spreadsheet-viewer');
    return $columns;
}
add_filter('manage_esv_spreadsheet_posts_columns', 'esv_spreadsheet_custom_columns');

// Populate column data
function esv_spreadsheet_custom_column_data($column, $post_id) {
    switch ($column) {
        case 'esv_table_id':
            echo esc_html(get_post_meta($post_id, 'esv_table_id', true));
            break;
        case 'esv_url':
            $url = esc_url(get_post_meta($post_id, 'esv_url', true));
            echo "<a href=\"" . esc_url($url) . "\" target=\"_blank\">" . 
            esc_html(strlen($url) > 40 ? substr($url, 0, 40) . '...' : $url) . 
            "</a>";            
            break;
        case 'esv_flattened':
            $path = get_post_meta($post_id, 'esv_flattened_path', true);
            if ($path && file_exists($path)) {
                echo esc_html(gmdate('Y-m-d H:i:s', filemtime($path)));
            } else {
                echo '<em>None</em>';
            }
            break;            
    }
}
add_action('manage_esv_spreadsheet_posts_custom_column', 'esv_spreadsheet_custom_column_data', 10, 2);
