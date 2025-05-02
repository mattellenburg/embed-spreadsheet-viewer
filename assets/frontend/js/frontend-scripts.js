jQuery(document).ready(function ($) {
    
    let currentColumnIndex = null;
    let currentTable = null;

    /*** --- CONTEXT MENU SETUP --- ***/
    $(document).on('contextmenu', '.esv-table th', function (e) {
        e.preventDefault();

        currentColumnIndex = $(this).index();
        currentTable = $(this).closest('table');

        const $menu = $('#esv-context-menu');

        $menu.css({
            display: 'block',
            position: 'absolute',
            top: e.pageY + 'px',
            left: e.pageX + 'px',
            zIndex: 10000
        });
    });

    $(document).on('click', function (e) {
        if (!$(e.target).closest('#esv-context-menu').length) {
            $('#esv-context-menu').hide();
        }
    });

    $(document).on('click', '#esv-context-menu li', function () {
        const action = $(this).data('action');
        if (!currentTable) return;

        switch (action) {
            case 'sort-asc':
                esvSortTable(currentTable, currentColumnIndex, true);
                break;
            case 'sort-desc':
                esvSortTable(currentTable, currentColumnIndex, false);
                break;
            case 'hide':
                esvHideColumn(currentTable, currentColumnIndex);
                break;
            case 'show-all':
                esvShowAllColumns(currentTable);
                break;
            case 'filter':
                esvFilterColumn(currentTable, currentColumnIndex);
                break;
            case 'clear-filters':
                esvClearFilters(currentTable);
                break;
        }

        $('#esv-context-menu').hide();
    });

    /*** --- SORT TABLE --- ***/
    function esvSortTable($table, colIndex, asc = true) {
        const $tbody = $table.find('tbody');
        const $rows = $tbody.find('tr').get();

        $rows.sort(function (a, b) {
            const A = $(a).children('td').eq(colIndex).text().toUpperCase();
            const B = $(b).children('td').eq(colIndex).text().toUpperCase();

            if ($.isNumeric(A) && $.isNumeric(B)) {
                return asc ? A - B : B - A;
            }
            return asc ? (A > B ? 1 : -1) : (A < B ? 1 : -1);
        });

        $.each($rows, function (index, row) {
            $tbody.append(row);
        });
    }

    /*** --- HIDE / SHOW COLUMNS --- ***/
    function esvHideColumn($table, colIndex) {
        $table.find('th, td').filter(function () {
            return $(this).index() === colIndex;
        }).hide();
    }

    function esvShowAllColumns($table) {
        $table.find('th, td').show();
    }

    /*** --- FILTERING --- ***/
    function esvFilterColumn($table, colIndex) {
        const filterValue = prompt("Enter filter text:");
        if (filterValue === null) return;

        $table.find('tbody tr').each(function () {
            const $td = $(this).children('td').eq(colIndex);
            if ($td.text().toLowerCase().indexOf(filterValue.toLowerCase()) === -1) {
                $(this).hide();
            } else {
                $(this).show();
            }
        });
    }

    /*** --- CLEAR ALL FILTERS --- ***/
    function esvClearFilters($table) {
        $table.find('tbody tr').show();
    }

    function esvInitPagination($table) {
        console.log('📜 Initializing pagination for table', $table);
    
        const $container = $table.closest('.esv-table-container');
        if ($container.length === 0) {
            console.log('❌ No container found for table');
            return;
        }
    
        console.log('✅ Found container:', $container);
    
        const $rows = $table.find('tbody tr');
        const totalRows = $rows.length;
    
        console.log('✅ Total rows:', totalRows);
    
        if (totalRows <= 10) {
            console.log('ℹ️ Only', totalRows, 'rows, skipping pagination');
            return; // No need to paginate
        }
    
        // Remove any existing pagination controls
        $container.find('.esv-pagination-controls').remove();
    
        const $pagination = $('<div class="esv-pagination-controls"></div>');
        const $select = $('<select class="esv-page-select"></select>');
        
        // Add page navigation controls
        const $pageControls = $('<div class="esv-page-navigation"></div>');
        $pageControls.append('<button class="esv-prev-page" disabled>&laquo; Prev</button>');
        $pageControls.append('<span class="esv-page-info">Page <input type="number" class="esv-current-page-input" value="1" min="1"> of <span class="esv-total-pages">1</span></span>');
        $pageControls.append('<button class="esv-next-page">Next &raquo;</button>');
    
        const options = [10];
        if (totalRows > 30) options.push(25);
        if (totalRows > 100) options.push(50);
        options.push('All');
    
        // Build dropdown
        $.each(options, function (i, val) {
            $select.append('<option value="' + val + '">' + val + '</option>');
        });
    
        $pagination.append('<label>Rows per page: </label>').append($select).append($pageControls);
        $container.append($pagination);
    
        console.log('✅ Pagination controls created and appended');
    
        // Initialize pagination state
        let currentPage = 1;
        let rowsPerPage = 10;
        
        function updatePageInfo() {
            const totalPages = rowsPerPage === 'All' ? 1 : Math.ceil(totalRows / rowsPerPage);
            $pagination.find('.esv-current-page').text(currentPage);
            $pagination.find('.esv-total-pages').text(totalPages);
            
            // Update button states
            $pagination.find('.esv-prev-page').prop('disabled', currentPage === 1);
            $pagination.find('.esv-next-page').prop('disabled', 
                currentPage === totalPages || rowsPerPage === 'All');
        }
    
        function paginate() {
            $rows.hide();
            
            if (rowsPerPage === 'All') {
                $rows.show();
                currentPage = 1;
            } else {
                const startIndex = (currentPage - 1) * rowsPerPage;
                const endIndex = startIndex + parseInt(rowsPerPage, 10);
                $rows.slice(startIndex, endIndex).show();
            }
            
            updatePageInfo();
        }
    
        // Initial pagination
        paginate();
    
        // Handle rows per page change
        $select.on('change', function () {
            rowsPerPage = $(this).val();
            currentPage = 1; // Reset to first page when changing rows per page
            paginate();
        });
        
        // Handle next page click
        $pagination.find('.esv-next-page').on('click', function() {
            if (rowsPerPage === 'All') return;
            
            const totalPages = Math.ceil(totalRows / rowsPerPage);
            if (currentPage < totalPages) {
                currentPage++;
                paginate();
            }
        });
        
        // Handle previous page click
        $pagination.find('.esv-prev-page').on('click', function() {
            if (currentPage > 1) {
                currentPage--;
                paginate();
            }
        });

        // Handle direct page input
        $pagination.find('.esv-current-page-input').on('keypress', function(e) {
            if (e.which === 13) { // Enter key
                e.preventDefault();
                
                const totalPages = rowsPerPage === 'All' ? 1 : Math.ceil(totalRows / rowsPerPage);
                let newPage = parseInt($(this).val(), 10);
                
                // Validate the input
                if (isNaN(newPage) || newPage < 1) {
                    newPage = 1;
                } else if (newPage > totalPages) {
                    newPage = totalPages;
                }
                
                // Update input value in case it was corrected
                $(this).val(newPage);
                
                // Only redraw if the page actually changed
                if (newPage !== currentPage) {
                    currentPage = newPage;
                    paginate();
                }
            }
        });
    }

    /*** --- INIT on document ready --- ***/
    $('.esv-table').each(function () {
        esvInitPagination($(this));
    });

    /*** --- INIT after AJAX Preview loads (dynamic) --- ***/
    $(document).on('preview-table-loaded', function (e, $newTable) {
        esvInitPagination($newTable);
    });

    $('.esv-refresh-btn').on('click', function (e) {
        e.preventDefault();

        const $btn = $(this);
        const url = $btn.data('url');
        const sheet = $btn.data('sheet');
        const postId = $btn.data('post');
        const $status = $('#esv-refresh-status-' + postId);

        $status.text('⏳ Refreshing...');

        $.post(esv_frontend.ajaxurl, {
            action: 'esv_retry_flatten',
            nonce: esv_frontend.nonce,
            url: url,
            worksheet: sheet,
            post_id: postId
        }, function (response) {
            if (response.success) {
                $status.html('✅ Updated! Reload the page to view.');
            } else {
                $status.html('❌ Error: ' + response.data.message);
            }
        });
    });

});
