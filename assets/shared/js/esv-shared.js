jQuery(document).ready(function($) {
    $(document).on('click', '.esv-retry-flatten, .esv-refresh-btn', function (e) {
        e.preventDefault();
    
        const $btn = $(this);
        const url = $btn.data('url');
        const sheet = $btn.data('sheet');
        const postId = $btn.data('post');
        const statusSelector = $btn.data('status-selector') || '#esv-refresh-status-' + postId;
        const $status = $(statusSelector);
    
        $status.text('⏳ Refreshing...');
    
        $.post(esv_shared.ajaxurl, {
            action: 'esv_retry_flatten',
            nonce: esv_shared.nonce,
            url: url,
            worksheet: sheet,
            post_id: postId
        }, function (response) {
            if (response.success) {
                $status.html('✅ Updated! Reload the page to view.');
            } else {
                const message = response.data && response.data.message 
                    ? response.data.message 
                    : 'Unknown error.';
                $status.html('❌ Error: ' + message);
            }
        });
    });     
});
