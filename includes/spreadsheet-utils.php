<?php

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

function esv_render_spreadsheet_table($table_id, $args = []) {
    $defaults = [
        'max_rows' => 10,
        'pagination' => true,
        'sticky_header' => true,
        'add_context_menu' => true,
    ];
    $args = wp_parse_args($args, $defaults);

    $table_id_safe = esc_attr($table_id); // Sanitize table ID

    $posts = get_posts([
        'post_type' => 'esv_spreadsheet',
        'posts_per_page' => 1,
        'meta_key' => 'esv_table_id',
        'meta_value' => $table_id,
    ]);

    if (empty($posts)) {
        return "<p><strong>Error:</strong> Spreadsheet with ID <strong>" . esc_html($table_id) . "</strong> not found.</p>";
    }

    $post = $posts[0];
    $meta = get_post_meta($post->ID);

    $max_rows = intval($args['max_rows']);

    $spreadsheet_name = get_the_title($post);
    $spreadsheet_name_safe = esc_html($spreadsheet_name); // Escape title

    $url = isset($meta['esv_url'][0]) ? esc_url($meta['esv_url'][0]) : '';
    $worksheet_name = isset($meta['esv_worksheet'][0]) ? sanitize_text_field($meta['esv_worksheet'][0]) : 'Sheet1';

    $header_row = intval($meta['esv_header_row'][0] ?? 1);
    $start_row = intval($meta['esv_start_row'][0] ?? ($header_row + 1));
    $end_row = intval($meta['esv_end_row'][0] ?? 0);
    $start_col = intval($meta['esv_start_col'][0] ?? 1);
    $end_col = intval($meta['esv_end_col'][0] ?? 0);

    $excluded_rows = esv_parse_ranges($meta['esv_excluded_rows'][0] ?? '');
    $excluded_cols = esv_parse_column_ranges($meta['esv_excluded_cols'][0] ?? '');
    $format_map = esv_parse_key_value_pairs($meta['esv_column_formats'][0] ?? '');
    $width_map = esv_parse_key_value_pairs($meta['esv_column_widths'][0] ?? '', 500);
    $custom_headers  = esv_parse_key_string_pairs($meta['esv_column_headers'][0] ?? '', '');

    if (!$url) {
        return "<p><strong>Error:</strong> Spreadsheet URL is missing.</p>";
    }

    $flattened_path = get_post_meta($post->ID, 'esv_flattened_path', true);

    if (empty($flattened_path) || !file_exists($flattened_path)) {
        return "<p><strong>Error:</strong> No flattened spreadsheet found. Please try saving again in the admin to generate a processed version.</p>";
    }

    try {
        $reader = IOFactory::createReaderForFile($flattened_path);
        $reader->setReadDataOnly(true);
        $reader->setReadEmptyCells(false);
        $spreadsheet = $reader->load($flattened_path);
    } catch (Exception $e) {
        return "<p><strong>Error:</strong> Failed to read flattened spreadsheet: " . esc_html($e->getMessage()) . "</p>";
    }

    $sheet = $spreadsheet->getSheetByName($worksheet_name);
    if (!$sheet) {
        return "<p><strong>Error:</strong> Worksheet '" . esc_html($worksheet_name) . "' not found.</p>";
    }

    $highestRow = $end_row ?: $sheet->getHighestRow();
    $highestColumnIndex = $end_col ?: Coordinate::columnIndexFromString($sheet->getHighestColumn());

    ob_start();

    if ($args['sticky_header']) {
        $show_download = get_post_meta($post->ID, 'esv_show_download_link', true);
        echo '<h3>' . $spreadsheet_name_safe;
        if ($show_download) {
            echo ' <a href="' . esc_url($url) . '" download>Download</a>';
        }
        echo '</h3>';
        echo '<div class="esv-meta"></div>';

        echo '<div class="esv-table-container">';
        echo '<p>Right click on a column to access sorting and filtering functionality. ' . ($max_rows !== 0 ? 'A maximum of ' . esc_html($highestRow) . ' rows are displayed. ' : '') . '</p>';

        $show_refresh = get_post_meta($post->ID, 'esv_show_refresh_button', true);
        if ($show_refresh) {
            echo '<button type="button" class="button esv-refresh-btn" 
                    data-url="' . esc_attr($url) . '" 
                    data-sheet="' . esc_attr($worksheet_name) . '" 
                    data-post="' . esc_attr($post->ID) . '" 
                    data-status-selector="#esv-refresh-status-' . esc_attr($post->ID) . '">
                    🔄 Refresh Spreadsheet
                    </button>';
        }

        echo '<i>Last Modified: ' . esc_html(date('Y-m-d H:i:s', filemtime($flattened_path))) . '</i>';
        echo '<div id="esv-refresh-status-' . esc_attr($post->ID) . '" class="esv-refresh-status"></div>';
    }

    echo '<table class="esv-table' . ($args['sticky_header'] ? ' esv-sticky-header' : '') . '" data-table-id="' . esc_attr($table_id) . '">';
    echo '<thead><tr>';

    for ($col = $start_col; $col <= $highestColumnIndex; $col++) {
        if (in_array($col, $excluded_cols)) continue;

        $header = $custom_headers[$col] ?? $sheet->getCell(Coordinate::stringFromColumnIndex($col) . $header_row)->getFormattedValue();
        echo '<th>' . esc_html($header) . '</th>';
    }
    echo '</tr></thead><tbody>';

    $rows_rendered = 0;
    foreach ($sheet->getRowIterator($start_row, $highestRow) as $rowObj) {
        $row = $rowObj->getRowIndex();
        if (in_array($row, $excluded_rows)) continue;

        echo '<tr>';
        foreach ($rowObj->getCellIterator() as $cell) {
            $colLetter = $cell->getColumn();
            $col = Coordinate::columnIndexFromString($colLetter);

            if ($col < $start_col || $col > $highestColumnIndex) continue;
            if (in_array($col, $excluded_cols)) continue;

            $raw_value = $cell->getValue();
            $hyperlink = $cell->hasHyperlink() ? $cell->getHyperlink()->getUrl() : null;
            $format_type = $format_map[$col] ?? 0;
            $display = esv_format_cell((string)$raw_value, $format_type);

            $escaped = esc_html($display);
            $max_width = intval($width_map[$col] ?? 500);

            echo '<td style="max-width:' . esc_attr($max_width) . 'px;">';
            echo '<div class="cell-text" title="' . $escaped . '">';
            echo $hyperlink ? '<a href="' . esc_url($hyperlink) . '" target="_blank">' . $escaped . '</a>' : $escaped;
            echo '</div></td>';
        }
        echo '</tr>';

        if ($max_rows > 0 && ++$rows_rendered >= $max_rows) {
            break;
        }
    }

    echo '</tbody></table>';

    if ($args['pagination']) {
        echo '<div class="esv-pagination-controls"></div>';
    }

    echo '</div>';

    if ($args['add_context_menu']) {
        echo '<div id="esv-context-menu" style="display:none;">
                <ul>
                    <li data-action="sort-asc">Sort Ascending</li>
                    <li data-action="sort-desc">Sort Descending</li>
                    <li data-action="filter">Filter...</li>
                    <li data-action="clear-filters">Clear Filters</li>
                    <li data-action="hide">Hide Column</li>
                    <li data-action="show-all">Show All Columns</li>
                </ul>
              </div>';
    }

    return ob_get_clean();
}

function esv_format_cell($value, $type) {
    switch (intval($type)) {
        case 1: return number_format((float)$value, 2); // Number
        case 2: // Date
            if (is_numeric($value)) {
                try {
                    return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($value)->format('Y-m-d');
                } catch (Exception $e) {
                    return $value;
                }
            }
            return date('Y-m-d', strtotime($value));
        case 3: return '$' . number_format((float)$value, 2); // Currency
        default: return $value;
    }
}

function esv_parse_ranges($input) {
    $result = [];
    $input = preg_replace('/\s+/', '', $input);
    $parts = explode(',', $input);
    foreach ($parts as $part) {
        if (preg_match('/^(\d+)-(\d+)$/', $part, $matches)) {
            $result = array_merge($result, range($matches[1], $matches[2]));
        } elseif (is_numeric($part)) {
            $result[] = intval($part);
        }
    }
    return array_unique($result);
}

function esv_parse_key_value_pairs($input, $default = 0) {
    $map = [];

    foreach (explode(',', $input) as $pair) {
        if (preg_match('/^([A-Z]+|\d+)=([\d.]+)$/i', trim($pair), $m)) {
            $colKey = is_numeric($m[1]) 
                ? intval($m[1]) 
                : \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString(strtoupper($m[1]));

            $map[$colKey] = $m[2];
        }
    }

    return $map;
}

function esv_parse_column_ranges($input) {
    $cols = esv_parse_ranges($input);
    preg_match_all('/([A-Z]+):([A-Z]+)/i', $input, $matches, PREG_SET_ORDER);
    foreach ($matches as $match) {
        $start = Coordinate::columnIndexFromString($match[1]);
        $end = Coordinate::columnIndexFromString($match[2]);
        $cols = array_merge($cols, range($start, $end));
    }
    return array_unique($cols);
}

function esv_parse_key_string_pairs($input) {
    $map = [];
    foreach (explode(',', $input) as $pair) {
        if (preg_match('/(\d+)=(.+)/', $pair, $m)) {
            $map[intval($m[1])] = trim($m[2]);
        }
    }
    return $map;
}

function esv_convert_dropbox_to_preview_url($url) {
    // If it's a Dropbox link, force ?dl=0
    if (strpos($url, 'dropbox.com') !== false) {
        // Replace existing dl=1 or dl=0
        $url = preg_replace('/[?&]dl=\d/', '', $url); // Remove existing dl
        $glue = strpos($url, '?') !== false ? '&' : '?';
        return $url . $glue . 'dl=0';
    }
    return $url;
}
