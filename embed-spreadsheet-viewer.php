<?php
/**
 * Plugin Name: Embed Spreadsheet Viewer
 * Description: Embed spreadsheet tables from shared URLs with customizable views.
 * Version: 1.0
 * Author: Matt Ellenburg
 * Text Domain: embed-spreadsheet-viewer
 * License: GPLv2 or later
 * License URI: https://www.gnu.org/licenses/gpl-2.0.html
 */

defined('ABSPATH') or die('No script kiddies please!');

// Includes
require_once plugin_dir_path(__FILE__) . 'includes/admin-menu.php';
require_once plugin_dir_path(__FILE__) . 'includes/ajax-preview.php';
require_once plugin_dir_path(__FILE__) . 'includes/spreadsheet-meta-boxes.php';
require_once plugin_dir_path(__FILE__) . 'includes/spreadsheet-post-type.php';
require_once plugin_dir_path(__FILE__) . 'includes/spreadsheet-shortcode.php';
require_once plugin_dir_path(__FILE__) . 'includes/spreadsheet-utils.php';
require_once plugin_dir_path(__FILE__) . 'includes/spreadsheet-values-extractor.php';
require_once plugin_dir_path(__FILE__) . 'vendor/autoload.php';

// Shared script (used by both admin + frontend)
function esv_enqueue_shared_scripts() {
    wp_enqueue_script('esv-shared-scripts', plugins_url('assets/shared/js/esv-shared.js', __FILE__), ['jquery'], null, true);

    wp_localize_script('esv-shared-scripts', 'esv_shared', [
        'ajaxurl' => admin_url('admin-ajax.php'),
        'nonce'   => wp_create_nonce('esv_ajax_preview'),
    ]);
}
add_action('admin_enqueue_scripts', 'esv_enqueue_shared_scripts');
add_action('wp_enqueue_scripts', 'esv_enqueue_shared_scripts');

// Admin assets
add_action('admin_enqueue_scripts', function () {
    $screen = get_current_screen();
    if ($screen && $screen->post_type === 'esv_spreadsheet') {
        // Main admin scripts
        wp_enqueue_script('esv-admin-scripts', plugins_url('assets/admin/js/admin-scripts.js', __FILE__), ['jquery'], null, true);

        // Excel values extractor script
        wp_enqueue_script('esv-values-extractor', plugins_url('assets/admin/js/spreadsheet-values-extractor.js', __FILE__), ['jquery'], null, true);

        wp_localize_script('esv-admin-scripts', 'esv_admin', [
            'ajaxurl' => admin_url('admin-ajax.php'),
            'nonce'   => wp_create_nonce('esv_ajax_preview'),
        ]);

        wp_localize_script('esv-values-extractor', 'MyPlugin', [
            'ajaxurl'        => admin_url('admin-ajax.php'),
            'security_nonce' => wp_create_nonce('my_plugin_nonce'),
        ]);
    }
});

// Frontend assets
add_action('wp_enqueue_scripts', function () {
    wp_enqueue_style('esv-view-styles', plugins_url('assets/frontend/css/frontend-styles.css', __FILE__));
    wp_enqueue_script('esv-view-scripts', plugins_url('assets/frontend/js/frontend-scripts.js', __FILE__), ['jquery'], null, true);
});

add_action('admin_notices', function () {
    $screen = get_current_screen();
    if ($screen && $screen->post_type === 'esv_spreadsheet') {
        echo '<div class="notice notice-info" style="padding:10px;">
            🍺 If this plugin saved you time, <a href="https://www.buymeacoffee.com/mattellenburg" target="_blank">buy me a beer</a> — cheers!
        </div>';
    }
});
