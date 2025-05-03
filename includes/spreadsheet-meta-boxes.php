<?php

// Add metabox
add_action('add_meta_boxes', 'esv_add_spreadsheet_meta_box');

function esv_add_spreadsheet_meta_box() {
    add_meta_box(
        'esv_spreadsheet_meta',
        __('Spreadsheet Settings', 'embed-spreadsheet-viewer'),
        'esv_render_spreadsheet_meta_box',
        'esv_spreadsheet',
        'normal',
        'high'
    );
}

// Render metabox UI
function esv_render_spreadsheet_meta_box($post) {
    $meta = get_post_meta($post->ID);
    $get = function ($key, $default = '') use ($meta) {
        return isset($meta[$key]) ? $meta[$key][0] : $default;
    };

    wp_nonce_field('esv_save_spreadsheet_meta', 'esv_spreadsheet_nonce');
    ?>

    <div id="esv-basic-fields">
        <table class="form-table">
            <tr>
                <th><label for="esv_table_id">Table ID</label></th>
                <td>
                    <input type="text" id="esv_table_id" name="esv_table_id" value="<?php echo esc_attr($get('esv_table_id')); ?>" class="regular-text" required>
                    <br><em>This is the ID used in the shortcode, e.g., [spreadsheet id="sample-data"].</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_url">Spreadsheet URL</label></th>
                <td>
                    <input type="url" id="esv_url" name="esv_url" value="<?php echo esc_url($get('esv_url')); ?>" class="regular-text" required>
                    <br><em>This URL must be publicly viewable and downloadable, e.g., Dropbox URLs should end with "=1".</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_worksheet">Worksheet Name</label></th>
                <td>
                    <input type="text" id="esv_worksheet" name="esv_worksheet" value="<?php echo esc_attr($get('esv_worksheet', 'Sheet1')); ?>" class="regular-text">
                    <br><em>Name of the worksheet that contains the data.</em>
                </td>
            </tr>
        </table>
    </div>

    <div id="esv-advanced-fields" style="<?php echo ($post->post_status === 'auto-draft') ? 'display:none;' : ''; ?>">
        <hr>
        <h3>Advanced Settings</h3>
        <table class="form-table">
            <tr>
                <th><label for="esv_header_row">Header Row</label></th>
                <td>
                    <input type="number" id="esv_header_row" name="esv_header_row" value="<?php echo esc_attr($get('esv_header_row', 1)); ?>" min="1" class="small-text"> <em>Row number that contains the column headings</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_start_row">Start Row</label></th>
                <td>
                    <input type="number" id="esv_start_row" name="esv_start_row" value="<?php echo esc_attr($get('esv_start_row', 2)); ?>" min="1" class="small-text"> <em>First data row</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_end_row">End Row</label></th>
                <td>
                    <input type="number" id="esv_end_row" name="esv_end_row" value="<?php echo esc_attr($get('esv_end_row')); ?>" min="1" class="small-text"> <em>Leave blank to autodetect</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_start_col">Start Column</label></th>
                <td>
                    <input type="number" id="esv_start_col" name="esv_start_col" value="<?php echo esc_attr($get('esv_start_col', 1)); ?>" min="1" class="small-text"> <em>First data column</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_end_col">End Column</label></th>
                <td>
                    <input type="number" id="esv_end_col" name="esv_end_col" value="<?php echo esc_attr($get('esv_end_col')); ?>" min="1" class="small-text"> <em>Leave blank to autodetect</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_excluded_rows">Exclude Rows</label></th>
                <td>
                    <input type="text" id="esv_excluded_rows" name="esv_excluded_rows" value="<?php echo esc_attr($get('esv_excluded_rows')); ?>" placeholder="e.g. 1,2,5-9" class="regular-text">
                    <br><em>Enter specific rows or ranges separated by commas</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_excluded_cols">Exclude Columns</label></th>
                <td>
                    <input type="text" id="esv_excluded_cols" name="esv_excluded_cols" value="<?php echo esc_attr($get('esv_excluded_cols')); ?>" placeholder="e.g. 1,2,5-8,M:O" class="regular-text">
                    <br><em>Enter specific column numbers/letters or ranges separated by commas</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_column_formats">Column Formats (non-text)</label></th>
                <td>
                    <input type="text" id="esv_column_formats" name="esv_column_formats" value="<?php echo esc_attr($get('esv_column_formats')); ?>" placeholder="e.g. 1=3,B=2,6=1" class="regular-text">
                    <br><em>Format Codes: 1 = Number, 2 = Date, 3 = Currency</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_column_widths">Column Maximum Widths</label></th>
                <td>
                    <input type="text" id="esv_column_widths" name="esv_column_widths" value="<?php echo esc_attr($get('esv_column_widths')); ?>" placeholder="e.g. 1=100,2=500" class="regular-text">
                    <br><em>Defaults to 500px max if not provided</em>
                </td>
            </tr>
            <tr>
                <th><label for="esv_column_headers">Column Headers</label></th>
                <td>
                    <input type="text" id="esv_column_headers" name="esv_column_headers" value="<?php echo esc_attr($get('esv_column_headers')); ?>" placeholder="e.g. 1=ID,B=Date" class="regular-text">
                    <br><em>Overrides values from defined header row</em>
                </td>
            </tr>
            <tr>
                <th><label>Cached Spreadsheet</label></th>
                <td>
                    <?php
                    $flattened_path = get_post_meta($post->ID, 'esv_flattened_path', true);
                    if ($flattened_path && file_exists($flattened_path)) {
                        echo '<p style="color:green;">✅ Spreadsheet generated.</p>';
                    } else {
                        echo '<p style="color:red;">⚠️ Spreadsheet not generated.</p>';
                    }
                    ?>
                    <button type="button" class="button esv-retry-flatten" 
                        data-url="<?php echo esc_attr($get('esv_url')); ?>" 
                        data-sheet="<?php echo esc_attr($get('esv_worksheet', 'Sheet1')); ?>" 
                        data-post="<?php echo esc_attr($post->ID); ?>" 
                        data-status-selector="#esv-flatten-status">
                        Recreate Spreadsheet
                    </button>
                    <div id="esv-flatten-status" style="margin-top:5px;"></div>
                </td>
            </tr>
            <tr>
                <th><label>Show Download Link</label></th>
                <td>
                    <input type="checkbox" id="esv_show_download_link" name="esv_show_download_link" value="1" <?php checked($get('esv_show_download_link'), '1'); ?>>
                    <em>Checking this box will let users download the worksheet.</em>
                </td>
            </tr>
            <tr>
                <th><label>Show Refresh Button</label></th>
                <td>
                    <input type="checkbox" id="esv_show_refresh_button" name="esv_show_refresh_button" value="1" <?php checked($get('esv_show_refresh_button'), '1'); ?>>
                    <em>Allowing end users to refresh spreadsheets may consume significant resources.</em>
                </td>
            </tr>
            <tr>
                <th></th>
                <td>
                    <button type="button" class="button esv-preview-button" data-mode="worksheet" data-table_id="<?php echo esc_attr($get('esv_table_id')); ?>" data-worksheet="<?php echo esc_attr($get('esv_worksheet', 'Sheet1')); ?>">Show Worksheet Preview</button>
                    <button type="button" class="button esv-preview-button" data-mode="table" data-table_id="<?php echo esc_attr($get('esv_table_id')); ?>" data-worksheet="<?php echo esc_attr($get('esv_worksheet', 'Sheet1')); ?>">Show Table Preview</button>
                    <br><em>You must publish the spreadsheet before you can see the preview.</em>
                </td>
            </tr>
            <tr>
                <th><label>Preview</label></th>
                <td>
                    <div id="esv-preview-container">
                        <div class="esv-preview-scroll">
                            <em>No preview loaded yet.</em>
                        </div>
                    </div>
                </td>
            </tr>
        </table>
    </div>
    <?php
}

// Save metabox data
function esv_save_spreadsheet_meta($post_id) {
    if (
        ! isset($_POST['esv_spreadsheet_nonce']) ||
        ! wp_verify_nonce(sanitize_text_field(wp_unslash($_POST['esv_spreadsheet_nonce'])), 'esv_save_spreadsheet_meta')
    ) {
        return;
    }

    if (defined('DOING_AUTOSAVE') && DOING_AUTOSAVE) return;
    if (!current_user_can('edit_post', $post_id)) return;

    $fields = [
        'esv_table_id',
        'esv_url',
        'esv_worksheet',
        'esv_header_row',
        'esv_start_row',
        'esv_end_row',
        'esv_start_col',
        'esv_end_col',
        'esv_excluded_rows',
        'esv_excluded_cols',
        'esv_column_formats',
        'esv_column_widths',
        'esv_column_headers',
        'esv_show_download_link',
        'esv_show_refresh_button',
    ];

    foreach ($fields as $field) {
        if (isset($_POST[$field])) {
            update_post_meta($post_id, $field, sanitize_text_field(wp_unslash($_POST[$field])));
        } else {
            delete_post_meta($post_id, $field);
        }
    }

    $url = isset($_POST['esv_url']) ? esc_url_raw(sanitize_text_field(wp_unslash($_POST['esv_url']))) : '';
    $sheet = isset($_POST['esv_worksheet']) ? sanitize_text_field(wp_unslash($_POST['esv_worksheet'])) : 'Sheet1';

    if (!empty($url)) {
        do_action('esv_after_spreadsheet_save', $url, $sheet, $post_id);
    }
}
add_action('save_post_esv_spreadsheet', 'esv_save_spreadsheet_meta');
