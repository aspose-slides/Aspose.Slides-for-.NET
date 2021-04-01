function shareApp(type) {
    if (['facebook', 'twitter', 'linkedin', 'feedback', 'bookmark'].indexOf(type) !== -1) {
        var gaEvent = function (action, category) {
            if (!category) {
                category = 'Social';
            }
            if ('ga' in window) {
                try {
                    var tracker = window.ga.getAll()[0];
                    if (tracker !== undefined) {
                        tracker.send('event', {
                            'eventCategory': category,
                            'eventAction': action
                        });
                    }
                } catch (err) { }
            }
        };
        var appURL = 'https://' + window.location.hostname + window.location.pathname;
        var title = document.title.replace('&', 'and');
        // Google Analytics event
        gaEvent(type.charAt(0).toUpperCase() + type.slice(1));

        // perform an action
        switch (type) {
            case 'facebook':
                var a = document.createElement('a');
                a.href = 'https://www.facebook.com/sharer/sharer.php?u=#' + encodeURI(appURL);
                a.setAttribute('target', '_blank');
                a.click();
                break;
            case 'twitter':
                var a = document.createElement('a');
                a.href = 'https://twitter.com/intent/tweet?text=' + encodeURI(title) + '&url=' + encodeURI(appURL);
                a.setAttribute('target', '_blank');
                a.click();
                break;
            case 'linkedin':
                var a = document.createElement('a');
                a.href = 'https://www.linkedin.com/sharing/share-offsite/?url=' + encodeURI(appURL);
                a.setAttribute('target', '_blank');
                a.click();
                break;
            case 'feedback':
                $('#feedbackModal').modal({
                    keyboard: true
                });
                break;
            case 'bookmark':
                $('#bookmarkModal').modal({
                    keyboard: true
                });
                break;
            default:
            // nothing
        }
    }
}

function sendFeedbackExtended(productName, appName, emailTo) {
    var text = $('#feedbackBody').val();
    if (text && !text.match(/^.s+$/)) {
        $('#feedbackModal').modal('hide');
        sendFeedback(text, productName, appName, emailTo);
    }
}

function sendFeedback(text, productName, appName, emailTo) {
    var msg = (typeof text === 'string' ? text : $('#feedbackText').val());
    if (!msg || msg.match(/^\s+$/) || msg.length > 1000) {
        return;
    }

    var data = {
        productName: productName,
        appName: appName,
        text: msg,
        emailTo: emailTo
    };

    if (!text) {
        if ('ga' in window) {
            try {
                var tracker = window.ga.getAll()[0];
                if (tracker !== undefined) {
                    tracker.send('event', {
                        'eventCategory': 'Social',
                        'eventAction': 'feedback-in-download'
                    });
                }
            } catch (e) { }
        }
    }

    $.ajax({
        method: "POST",
        url: '/common/sendfeedback',
        data: data,
        dataType: "json",
        success: (data) => {
            showToast(data.message);
            $('#feedback').hide();
        },
        error: (data) => {
            showAlert(data.responseJSON.message);
        }
    });
}

function showToast(msg) {
    var toast = $('.toast');
    if (toast.length <= 0) {
        toast = $('<div class="toast" role="alert" aria-live="assertive" aria-atomic="true" data-delay="2000" style="position: absolute; top: 5rem; right: 3rem;"><div class="toast-body">Hello</div></div>');
    }
    toast.find('.toast-body').text(msg);
    $('body').append(toast);
    toast.toast('show');
}


$(document).ready(function() {

    //  modal
    $('#bookmarkModal').on('show.bs.modal',
        function(e) {
            $('#bookmarkModal').css('display', 'flex');
            $('#bookmarkModal').on('keydown',
                function(evt) {
                    if ((evt.metaKey || evt.ctrlKey) && String.fromCharCode(evt.which).toLowerCase() === 'd') {
                        $('#bookmarkModal').modal('hide');
                    }
                });
        });
    $('#bookmarkModal').on('hidden.bs.modal',
        function(e) {
            $('#bookmarkModal').off('keydown');
        });

    // send feedback modal
    $('#feedbackModal').on('show.bs.modal',
        function(e) {
            $('#feedbackModal').css('display', 'flex');
        });
    $('#feedbackModal').on('shown.bs.modal',
        function() {
            $('#feedbackBody').focus();
        });
});
