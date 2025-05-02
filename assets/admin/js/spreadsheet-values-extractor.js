/**
 * Excel Values Extractor JavaScript
 * 
 * Functions to handle the extraction of values from Excel files
 */

// Ensure the object exists
var esv = esv || {};

/**
 * Excel Values Extractor module
 */
esv.ExcelValuesExtractor = (function($) {
    /**
     * Process an Excel file to extract values
     * 
     * @param {string} excelUrl URL of the Excel file
     * @param {string} sheetName Name of the worksheet
     * @param {string} prefix Optional prefix for file naming
     * @param {function} successCallback Function to call on success
     * @param {function} errorCallback Function to call on error
     */
    function processExcelFile(excelUrl, sheetName, prefix, successCallback, errorCallback) {
        // Show loading indicator or message
        showLoading(true);
        
        // Make AJAX request to process the file
        $.ajax({
            url: ajaxurl,
            type: 'POST',
            data: {
                action: 'process_excel_values',
                excel_url: excelUrl,
                sheet_name: sheetName,
                prefix: prefix || '', // Optional prefix
                nonce: MyPlugin.security_nonce // Assume this is set elsewhere in your admin.js
            },
            success: function(response) {
                showLoading(false);
                
                if (response.success) {
                    if (typeof successCallback === 'function') {
                        successCallback(response.data);
                    } else {
                        // Default success handler
                        handleSuccess(response.data);
                    }
                } else {
                    if (typeof errorCallback === 'function') {
                        errorCallback(response.data.message);
                    } else {
                        // Default error handler
                        handleError(response.data.message);
                    }
                }
            },
            error: function(xhr, status, error) {
                showLoading(false);
                
                if (typeof errorCallback === 'function') {
                    errorCallback('AJAX error: ' + error);
                } else {
                    // Default error handler
                    handleError('AJAX error: ' + error);
                }
            }
        });
    }
    
    /**
     * Show or hide loading indicator
     * 
     * @param {boolean} isLoading Whether loading is in progress
     */
    function showLoading(isLoading) {
        if (isLoading) {
            // Add your loading indicator here
            $('#excel-processing-status').html('<p>Processing Excel file... This may take a moment.</p>');
        } else {
            // Remove loading indicator
            $('#excel-processing-status').html('');
        }
    }
    
    /**
     * Default success handler
     * 
     * @param {object} data Response data
     */
    function handleSuccess(data) {
        // Display success message
        $('#excel-processing-status').html('<p>Excel file processed successfully!</p>');
        
        // Show download link
        $('#excel-results').show();
        $('#excel-original-file').text(data.original_file);
        $('#excel-values-file').text(data.values_file);
        $('#excel-download-link').attr('href', data.download_url);
    }
    
    /**
     * Default error handler
     * 
     * @param {string} message Error message
     */
    function handleError(message) {
        $('#excel-processing-status').html('<p class="error">Error: ' + message + '</p>');
    }
    
    /**
     * Add event listeners
     */
    function init() {
        // Example of how to bind to a button
        $(document).on('click', '#process-excel-btn', function(e) {
            e.preventDefault();
            
            // You would get these values from your form or other source
            var excelUrl = $('#excel-url').val();
            var sheetName = $('#sheet-name').val();
            
            processExcelFile(excelUrl, sheetName, '', handleSuccess, handleError);
        });
    }
    
    // Return public methods
    return {
        init: init,
        processExcelFile: processExcelFile
    };
})(jQuery);

// Initialize when document is ready
jQuery(document).ready(function() {
    // Initialize the Excel Values Extractor module
    esv.ExcelValuesExtractor.init();
});