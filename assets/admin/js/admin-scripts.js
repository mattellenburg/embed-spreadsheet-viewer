jQuery(document).ready(function ($) {

    function triggerPreview(mode) {
        const $container = $('#esv-preview-container');

        $container.html('<div class="esv-spinner">Loading preview...</div>');

        const data = {
            action: 'esv_preview_spreadsheet',
            security: esv_admin.nonce,
            url: $('#esv_url').val(),
            worksheet: $('#esv_worksheet').val(),
            mode: mode
        };

        if (mode === 'table') {
            Object.assign(data, {
                header_row: $('#esv_header_row').val(),
                start_row: $('#esv_start_row').val(),
                start_col: $('#esv_start_col').val(),
                excluded_rows: $('#esv_excluded_rows').val(),
                excluded_cols: $('#esv_excluded_cols').val()
            });
        }

        $.post(esv_admin.ajaxurl, data, function (response) {
            if (response.success && response.data) {
                $container.fadeOut(200, function () {
                    $(this).html(response.data).fadeIn(200);
                    $('html, body').animate({
                        scrollTop: $container.offset().top - 100
                    }, 600);
                });
            } else {
                $container.html('<p>Failed to load preview.</p>');
            }
        });
    }

    $('.esv-preview-button').on('click', function () {
        var $btn = $(this);
        var mode = $btn.data('mode');
        var worksheet = $btn.data('worksheet');
        var table_id = $btn.data('table_id');
    
        $('#esv-preview-container .esv-preview-scroll').html('<em>Loading preview...</em>');
    
        $.post(esv_admin.ajaxurl, {
            action: 'esv_preview_spreadsheet',
            security: esv_admin.nonce,
            mode: mode,
            worksheet: worksheet,
            table_id: table_id
        }, function (response) {
            if (response.success) {
                $('#esv-preview-container .esv-preview-scroll').html(response.data);
            } else {
                $('#esv-preview-container .esv-preview-scroll').html('<em>Error: ' + response.data + '</em>');
            }
        });
    });
    
    /*** --- Auto-focus & Scroll after Save --- ***/
    const $advancedFields = $('#esv-advanced-fields');

    if ($('#post_status').val() === 'auto-draft') {
        $advancedFields.hide();
    }

    $('#publish').on('click', function () {
        setTimeout(function () {
            $advancedFields.fadeIn('slow', function () {
                $('html, body').animate({
                    scrollTop: $container.offset().top - 100
                }, 600);                

                setTimeout(function () {
                    const $focusField = $('#esv_header_row');
                    $focusField.focus().addClass('esv-glow');

                    setTimeout(function () {
                        $focusField.removeClass('esv-glow');
                    }, 1500);
                }, 700);
            });
        }, 500);
    });

    if ($('#post_status').val() !== 'auto-draft') {
        $advancedFields.show();
    }

    // Hide Advanced on page load if not published
    if ($('#post_status').val() === 'auto-draft') {
        $advancedFields.hide();
    }

    // Watch for Save/Publish click
    $('#publish').on('click', function() {
        setTimeout(function() {
            $advancedFields.fadeIn('slow');
        }, 500); // slight delay after click to allow WordPress to update the post_status
    });

    // Also auto-show if user reloads page and it's now published
    if ($('#post_status').val() !== 'auto-draft') {
        $advancedFields.show();
    }

});
