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

    $posts = get_posts([
        'post_type'      => 'esv_spreadsheet',
        'posts_per_page' => 1,
        'meta_query'     => [
            [
                'key'     => 'esv_table_id',
                'value'   => $table_id,
                'compare' => '='
            ]
        ]
    ]);

    if (empty($posts)) {
        return "<p><strong>Error:</strong> Spreadsheet with ID <strong>" . esc_html($table_id) . "</strong> not found.</p>";
    }

    $post = $posts[0];
    $meta = get_post_meta($post->ID);

    $max_rows = intval($args['max_rows']);

    $spreadsheet_name = get_the_title($post);
    $spreadsheet_name_safe = esc_html($spreadsheet_name);

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
    $custom_headers  = esv_parse_key_string_pairs($meta['esv_column_headers'][0] ?? '');

    if (!$url) {
        return "<p><strong>Error:</strong> Spreadsheet URL is missing.</p>";
    }

    $original_path = get_post_meta($post->ID, 'esv_original_path', true);
    $flattened_path = get_post_meta($post->ID, 'esv_flattened_path', true);

    if (empty($flattened_path) || !file_exists($flattened_path)) {
        return "<p><strong>Error:</strong> No flattened spreadsheet found. Please try saving again in the admin to generate a processed version.</p>";
    }

    try {
        $reader = IOFactory::createReaderForFile($flattened_path);
        $reader->setReadDataOnly(false);
        $spreadsheet = $reader->load($flattened_path);
    } catch (Exception $e) {
        return "<p><strong>Error:</strong> Failed to read spreadsheet: " . esc_html($e->getMessage()) . "</p>";
    }

    $sheet = $spreadsheet->getSheetByName($worksheet_name);
    if (!$sheet) {
        return "<p><strong>Error:</strong> Worksheet '" . esc_html($worksheet_name) . "' not found.</p>";
    }

    $highestRow = $end_row ?: $sheet->getHighestRow();
    $highestColumnIndex = $end_col ?: Coordinate::columnIndexFromString($sheet->getHighestColumn());

    $hyperlink_map = [];
    if (!empty($original_path) && file_exists($original_path)) {
        try {
            $originalReader = IOFactory::createReaderForFile($original_path);
            $originalSpreadsheet = $originalReader->load($original_path);
            $originalSheet = $originalSpreadsheet->getSheetByName($worksheet_name);

            foreach ($originalSheet->getRowIterator($start_row, $highestRow) as $rowObj) {
                $rowIndex = $rowObj->getRowIndex();
                foreach ($rowObj->getCellIterator() as $cell) {
                    $colLetter = $cell->getColumn();
                    if ($cell->hasHyperlink()) {
                        $hyperlink_map[$rowIndex][$colLetter] = $cell->getHyperlink()->getUrl();
                    } else {
                        if ($cell->getDataType() === \PhpOffice\PhpSpreadsheet\Cell\DataType::TYPE_FORMULA) {
                            $formula = $cell->getValue();
                            if (preg_match('/HYPERLINK\(["\'](.*?)["\']/', $formula, $matches)) {
                                $hyperlink_map[$rowIndex][$colLetter] = $matches[1];
                            }
                        }
                    }
                }
            }
        } catch (Exception $e) {
            return "<p><strong>Error:</strong> Failed to load original file for hyperlinks: " . $e->getMessage() . "</p>";
        }
    }

    ob_start();

    if ($args['sticky_header']) {
        $show_download = get_post_meta($post->ID, 'esv_show_download_link', true);
        echo '<h3>' . esc_attr($spreadsheet_name_safe);
        if ($show_download) {
            echo ' <a href="' . esc_url($url) . '" download>Download</a>';
        }
        echo '</h3>';
        echo '<div class="esv-meta"></div>';

        echo '<div class="esv-table-container">';
        echo '<p>Right click on a column to access sorting and filtering functionality.</p>';

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

        echo '&nbsp;<i>Last Modified: ' . esc_html(gmdate('Y-m-d H:i:s', filemtime($flattened_path))) . '</i>';
        echo '<div id="esv-refresh-status-' . esc_attr($post->ID) . '" class="esv-refresh-status"></div>';
    }

    echo '<table class="esv-table' . ($args['sticky_header'] ? ' esv-sticky-header' : '') . '" data-table-id="' . esc_attr($table_id) . '">';
    echo '<thead><tr>';

    for ($col = $start_col; $col <= $highestColumnIndex; $col++) {
        $colLetter = Coordinate::stringFromColumnIndex($col);

        if (in_array($col, $excluded_cols, true) || in_array($colLetter, $excluded_cols, true)) continue;

        $header = $custom_headers[$col] ?? $custom_headers[$colLetter] ?? $sheet->getCell($colLetter . $header_row)->getFormattedValue();

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
            if (in_array($col, $excluded_cols, true) || in_array($colLetter, $excluded_cols, true)) continue;

            $raw_value = $cell->getValue();
            $format_type = $format_map[$col] ?? $format_map[$colLetter] ?? 0;
            $display = esv_format_cell((string)$raw_value, $format_type);
            $max_width = intval($width_map[$col] ?? $width_map[$colLetter] ?? 500);

            $hyperlink = $hyperlink_map[$row][$colLetter] ?? null;

            echo '<td style="max-width:' . esc_attr($max_width) . 'px;"><div class="cell-text" title="' . esc_attr($display) . '">';
            if ($hyperlink) {
                echo '<a href="' . esc_url($hyperlink) . '" target="_blank">' . esc_html($display) . '</a>';
            } else {
                echo esc_html($display);
            }
            echo '</div></td>';
        }
        echo '</tr>';

        if ($max_rows > 0 && ++$rows_rendered >= $max_rows) {
            break;
        }
    }

    echo '</tbody></table>';

    if ($args['pagination'] && $args['sticky_header']) {
        echo '<div class="esv-pagination-controls"></div>';
    }

    echo '</div>';

    if ($args['add_context_menu'] && $args['sticky_header']) {
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

// Updated helpers to handle both numbers and letters
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

function esv_parse_column_ranges($input) {
    $cols = [];
    $parts = explode(',', preg_replace('/\s+/', '', $input));
    foreach ($parts as $part) {
        if (ctype_alpha($part)) {
            $index = Coordinate::columnIndexFromString(strtoupper($part));
            $cols[] = $index;
            $cols[] = strtoupper($part);
        } elseif (is_numeric($part)) {
            $cols[] = intval($part);
            $cols[] = Coordinate::stringFromColumnIndex(intval($part));
        }
    }
    preg_match_all('/([A-Z]+):([A-Z]+)/i', $input, $matches, PREG_SET_ORDER);
    foreach ($matches as $match) {
        $start = Coordinate::columnIndexFromString($match[1]);
        $end = Coordinate::columnIndexFromString($match[2]);
        for ($i = $start; $i <= $end; $i++) {
            $cols[] = $i;
            $cols[] = Coordinate::stringFromColumnIndex($i);
        }
    }
    return array_unique($cols);
}

function esv_parse_key_value_pairs($input, $default = 0) {
    $map = [];
    foreach (explode(',', $input) as $pair) {
        if (preg_match('/^([A-Z]+|\d+)=([\d.]+)$/i', trim($pair), $m)) {
            $colKey = $m[1];
            $value = $m[2];

            $index = ctype_alpha($colKey)
                ? Coordinate::columnIndexFromString(strtoupper($colKey))
                : intval($colKey);

            $letter = Coordinate::stringFromColumnIndex($index);

            $map[$index] = $value;
            $map[$letter] = $value;
        }
    }
    return $map;
}

function esv_parse_key_string_pairs($input) {
    $map = [];
    foreach (explode(',', $input) as $pair) {
        if (preg_match('/([A-Z]+|\d+)=(.+)/', $pair, $m)) {
            $colKey = trim($m[1]);
            $label = trim($m[2]);

            $index = ctype_alpha($colKey)
                ? Coordinate::columnIndexFromString(strtoupper($colKey))
                : intval($colKey);

            $letter = Coordinate::stringFromColumnIndex($index);

            $map[$index] = $label;
            $map[$letter] = $label;
        }
    }
    return $map;
}

// Helper functions (unchanged)
function esv_format_cell($value, $type) {
    switch (intval($type)) {
        case 1: return number_format((float)$value, 2);
        case 2:
            if (is_numeric($value)) {
                try {
                    return \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($value)->format('Y-m-d');
                } catch (Exception $e) {
                    return $value;
                }
            }
            return gmdate('Y-m-d', strtotime($value));
        case 3: return '$' . number_format((float)$value, 2);
        default: return $value;
    }
}
