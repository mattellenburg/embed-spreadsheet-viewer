<?php

function esv_admin_menu() {
    add_menu_page('Spreadsheets', 'Spreadsheets', 'manage_options', 'esv_spreadsheet_menu', 'esv_render_list_page', 'dashicons-media-spreadsheet', 6);
    add_submenu_page('esv_spreadsheet_menu', 'Add Spreadsheet', 'Add Spreadsheet', 'manage_options', 'post-new.php?post_type=esv_spreadsheet');
}
add_action('admin_menu', 'esv_admin_menu');

// Enqueue admin assets only on Spreadsheet post type screen
add_action('admin_enqueue_scripts', function ($hook) {
    $screen = get_current_screen();

    if ($screen && $screen->post_type === 'esv_spreadsheet') {
        $base_url = plugin_dir_url(__DIR__);

        wp_enqueue_style('esv-admin-styles', $base_url . 'assets/admin/css/admin-styles.css', array(), ESV_VERSION);
        wp_enqueue_script('esv-admin-scripts', $base_url . 'assets/admin/js/admin-scripts.js', ['jquery'], ESV_VERSION, true);

        wp_localize_script('esv-admin-scripts', 'esv_admin', [
            'nonce' => wp_create_nonce('esv_ajax_preview'),
        ]);
    }
});

/**
 * AJAX handler for processing Excel files
 */
function process_excel_values_ajax() {
    // Verify nonce
    if (!isset($_POST['nonce']) || !wp_verify_nonce(sanitize_text_field(wp_unslash($_POST['nonce']), 'my_plugin_nonce'))) {
        wp_send_json_error(array('message' => 'Security check failed.'));
    }
    
    // Check capabilities
    if (!current_user_can('edit_posts')) {
        wp_send_json_error(array('message' => 'You do not have permission to perform this action.'));
    }
    
    // Check required parameters
    if (empty($_POST['excel_url']) || empty($_POST['sheet_name'])) {
        wp_send_json_error(array('message' => 'Excel URL and sheet name are required.'));
    }

    // Properly unslash and sanitize excel_url
    $excel_url = esc_url(sanitize_text_field(wp_unslash($_POST['excel_url'])));

    // Properly unslash and sanitize sheet_name
    $sheet_name = sanitize_text_field(wp_unslash($_POST['sheet_name']));

    // Properly unslash and sanitize prefix if it exists
    $prefix = '';
    if (isset($_POST['prefix'])) {
        $prefix = sanitize_text_field(wp_unslash($_POST['prefix']));
    }

    // Process the Excel file
    $result = extract_excel_values($excel_url, $sheet_name, $prefix);
    
    if (is_wp_error($result)) {
        wp_send_json_error(array('message' => $result->get_error_message()));
    } else {
        wp_send_json_success($result);
    }
}
add_action('wp_ajax_process_excel_values', 'process_excel_values_ajax');