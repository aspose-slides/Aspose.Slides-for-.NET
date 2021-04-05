function isEmailValid(email) {
    var re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
    return re.test(String(email).toLowerCase());
}

function validateEmailAndAlert(email) {
    if (!isEmailValid(email)) {
        showAlert(window.Resources["ValidateEmailMessage"]);
        return false;
    }
    else
        hideAlert();

    return true;
}

function sendEmail(data, textStatus, xhr) {
    let email = $('#EmailToInput').val();
    if (!validateEmailAndAlert(email))
        return;
    
    let nodes = document.getElementsByName("DownloadUrlInput");
    let urls = [];

    nodes.forEach(node => {
        urls.push(node.value);
    });

    $('#sendEmailButton').addClass("hidden");

    $.ajax({
        method: "POST",
        url: `/slides/common/SendUrlToEmail/${Method}`,
        data: { urls: urls, email: email },
        dataType: "json",
        beforeSend: showLoader,
        success: sendEmailSuccess,
        complete: completeSendEmail,
        error: handleError
    });
}

function completeSendEmail() {
    $('#sendEmailButton').removeClass("hidden");
    hideLoader();
}

function sendEmailSuccess(data, textStatus, xhr) {
    showInfo(data.message);
    SEONavigate('email');
}

$(document).ready(function () {
    $('#sendEmailButton').click(sendEmail);
});