$.extend($.expr[':'], {
    isEmpty: function (e) {
        return e.value === '';
    }
});

var getAbsoluteUrl = (function () {
    var a;
    return function (url) {
        if (!a) a = document.createElement('a');
        a.href = url;

        return a.href;
    };
})();

window.onWorkSuccess = function () {
};

window.workSuccess = function (data, textStatus, xhr) {
    hideLoader();

    if (Array.isArray(data)) {
        const uploaded = data.filter(f => f.IsSuccess).map(f => f);
        let downloadPlaceHolder;

        if (uploaded.length !== 0) {
            if (window.nextStage) {
                window.nextStage(data);
                return;
            }

            $('.UploadPlaceHolder').addClass("hidden");
            $('#WorkPlaceHolder').addClass("hidden");
            $('.add-property').toggle(false);
            $('.saveas').toggle(false);
            $('#chartTabs').toggle(false);
            
            downloadPlaceHolder = document.getElementById("DownloadPlaceHolder");
            downloadPlaceHolder.classList.remove("hidden");

            $('#OtherApps').removeClass("hidden");
            SEONavigate('result');
        }

        let lastDownloadButtonSeparator;
        let sendEmailButtonParent;
        let sendEmailButtonParentDiv;
        let downloadButtonText;
        let downloadButtonId;
        let downloadUrlInputId;
        let viewerplaceholderDiv;
        let lastViewerButtonSeparator;
        let viewerButtonText;
        let viewerButtonId;

        let editorplaceholderDiv;
        let lastEditorButtonSeparator;
        let editorButtonText;
        let editorButtonId;

        data.forEach(upFile => {
            if (upFile.IsSuccess) {

                let url =
                    `${APIBasePath}api/Common/DownloadFile/${upFile.id}?file=${encodeURIComponent(upFile.FileName)}`;
                url = getAbsoluteUrl(url);

                let index = data.indexOf(upFile);
                let downloadButton;
                let downloadUrlInput;
                let viewerButton;
                let editorButton;

                if (index === 0) {
                    downloadUrlInput = document.getElementById('DownloadUrlInputHidden');
                    downloadButton = document.getElementById('DownloadButton');

                    if (uploaded.length > 1) {
                        lastDownloadButtonSeparator = downloadButton.parentElement.nextElementSibling;
                        sendEmailButtonParent = document.getElementById("sendEmailButton").parentElement;
                        sendEmailButtonParentDiv = sendEmailButtonParent.parentElement;
                        downloadButtonText = downloadButton.innerHTML;
                        downloadButtonId = downloadButton.getAttribute('id');
                        downloadUrlInputId = downloadUrlInput.getAttribute("id");
                        // text
                        downloadButton.innerHTML = `${downloadButtonText} ${upFile.FileName}`;
                        // id
                        downloadButton.setAttribute('id', `${downloadButtonId}_${upFile.FileName}`);
                        downloadUrlInput.setAttribute("id", `${downloadUrlInputId}_${upFile.FileName}`);
                    }

                    // url
                    downloadButton.setAttribute('href', url);
                    downloadUrlInput.setAttribute("value", url);
                } else {
                    let downloadButtonContent = `${downloadButtonText} ${upFile.FileName}`;
                    let downloadButtonIdContent = `${downloadButtonId}_${upFile.FileName}`;
                    downloadButton = getDownloadButton(downloadButtonContent, downloadButtonIdContent, url);

                    let filesuccessDiv = document.getElementsByClassName("filesuccess")[0];
                    filesuccessDiv.insertBefore(downloadButton, lastDownloadButtonSeparator);

                    let downloadUrlInputIdContent = `${downloadUrlInputId}_${upFile.FileName}`;
                    downloadUrlInput = getDownloadUrlInput(downloadUrlInputIdContent, url);
                    sendEmailButtonParentDiv.insertBefore(downloadUrlInput, sendEmailButtonParent);
                }

                if (window.AllowedViewAfterProcesing && window.AllowedViewAfterProcesing(upFile)) {
                    /*
                    Query String is used here because escaped filename is not allowed in the Request Path and we will get
                    "A potentially dangerous Request.Path value was detected from the client."
                    see here: https://www.hanselman.com/blog/ExperimentsInWackinessAllowingPercentsAnglebracketsAndOtherNaughtyThingsInTheASPNETIISRequestURL.aspx
                    */
                    let viewUrl =
                        `/slides/storage/view/?folder=${upFile.id}&fileName=${encodeURIComponent(upFile.FileName)}`;
                    viewUrl = getAbsoluteUrl(viewUrl);
                    let editUrl =
                        `/slides/storage/edit/?copy=processed&folder=${upFile.id}&fileName=${encodeURIComponent(
                            upFile.FileName)}`;
                    editUrl = getAbsoluteUrl(editUrl);

                    if (index === 0) {

                        viewerplaceholderDiv = document.getElementById("viewerplaceholder");
                        viewerplaceholderDiv.classList.remove("hidden");
                        viewerButton = document.getElementById("ViewerButton");

                        editorplaceholderDiv = document.getElementById("editorPlaceholder");
                        editorplaceholderDiv.classList.remove("hidden");
                        editorButton = document.getElementById("EditorButton");

                        if (uploaded.length > 1) {
                            lastViewerButtonSeparator = viewerButton.parentElement.nextElementSibling;
                            viewerButtonText = viewerButton.innerHTML;
                            viewerButtonId = viewerButton.getAttribute('id');
                            // text
                            viewerButton.innerHTML = `${viewerButtonText} ${upFile.FileName}`;
                            // id
                            viewerButton.setAttribute('id', `${viewerButtonId}_${upFile.FileName}`);

                            lastEditorButtonSeparator = editorButton.parentElement.nextElementSibling;
                            editorButtonText = editorButton.innerHTML;
                            editorButtonId = editorButton.getAttribute('id');
                            // text
                            editorButton.innerHTML = `${editorButtonText} ${upFile.FileName}`;
                            // id
                            editorButton.setAttribute('id', `${editorButtonId}_${upFile.FileName}`);
                        }

                        // url
                        viewerButton.setAttribute("href", viewUrl);
                        editorButton.setAttribute("href", editUrl);
                    } else {
                        let viewerButtonContent = `${viewerButtonText} ${upFile.FileName}`;
                        let viewerButtonIdContent = `${viewerButtonId}_${upFile.FileName}`;
                        viewerButton = getViewerButton(viewerButtonContent, viewerButtonIdContent, viewUrl);

                        viewerplaceholderDiv.insertBefore(viewerButton, lastViewerButtonSeparator);

                        let editorButtonContent = `${editorButtonText} ${upFile.FileName}`;
                        let editorButtonIdContent = `${editorButtonId}_${upFile.FileName}`;
                        editorButton = getEditorButton(editorButtonContent, editorButtonIdContent, editUrl);

                        editorplaceholderDiv.insertBefore(editorButton, lastEditorButtonSeparator);
                    }
                }

                onWorkSuccess();
            } else {
                dataErrorAlert(upFile);
                if (window.onShowProcessingButton) {
                    window.onShowProcessingButton();
                }
            }
        });
    }
    else {
        if (data.IsSuccess) {

            if (window.nextStage) {
                window.nextStage(data);
                return;
            }

            $('.UploadPlaceHolder').addClass("hidden");
            $('#WorkPlaceHolder').addClass("hidden");
            $('.add-property').toggle(false);
            $('.saveas').toggle(false);
            $('#chartTabs').toggle(false);
            $('#DownloadPlaceHolder').removeClass("hidden");
            $('#OtherApps').removeClass("hidden");

            let url = `${APIBasePath}api/Common/DownloadFile/${data.id}?file=${encodeURIComponent(data.FileName)}`;
            url = getAbsoluteUrl(url);

            $('#DownloadUrlInputHidden').val(url);
            $('#DownloadButton').attr("href", url);

            SEONavigate('result');

            if (window.AllowedViewAfterProcesing && window.AllowedViewAfterProcesing(data)) {
                /*
                Query String is used here because escaped filename is not allowed in the Request Path and we will get
                "A potentially dangerous Request.Path value was detected from the client."
                see here: https://www.hanselman.com/blog/ExperimentsInWackinessAllowingPercentsAnglebracketsAndOtherNaughtyThingsInTheASPNETIISRequestURL.aspx
                */
                let viewurl = `/slides/storage/view/?folder=${data.id}&fileName=${encodeURIComponent(data.FileName)}`;
                viewurl = getAbsoluteUrl(viewurl);
                $('#ViewerButton').attr("href", viewurl);
                $('#viewerplaceholder').removeClass("hidden");

                let editUrl = `/slides/storage/edit/?copy=processed&folder=${data.id}&fileName=${encodeURIComponent(data.FileName)}`;
                editUrl = getAbsoluteUrl(editUrl);
                $('#EditorButton').attr("href", editUrl);
                $('#editorPlaceholder').removeClass("hidden");
            }

            onWorkSuccess();
        }
        else {
            dataErrorAlert(data);
            if (window.onShowProcessingButton) {
                window.onShowProcessingButton();
            }
        }
    }
}

function getDownloadButton(content, idContent, url) {
    let span = document.createElement("span");
    span.setAttribute("class", "downloadbtn convertbtn full-width");

    let a = document.createElement("a");
    a.setAttribute("class", "btn btn-success btn-lg btn-block");

    let i = document.createElement("i");
    i.setAttribute("class", "fa fa-download");

    a.appendChild(i);
    span.appendChild(a);

    // text
    a.innerHTML = content;
    // id
    a.setAttribute('id', idContent);
    // url
    a.setAttribute('href', url);

    return span;
}

function getDownloadUrlInput(idContent, url) {
    let downloadUrlInput = document.createElement("input");
    downloadUrlInput.setAttribute("type", "hidden");
    downloadUrlInput.setAttribute("name", "DownloadUrlInput");
    downloadUrlInput.setAttribute("id", idContent);
    downloadUrlInput.setAttribute("value", url);

    return downloadUrlInput;
}

function getViewerButton(content, idContent, url) {
    let span = document.createElement("span");
    span.setAttribute("class", "viewerbtn full-width");

    let a = document.createElement("a");
    a.setAttribute("class", "btn btn-success btn-lg btn-block");
    a.setAttribute("target", "_blank");

    let i = document.createElement("i");
    i.setAttribute("class", "fa fa-eye");
    a.appendChild(i);
    span.appendChild(a);

    // text
    a.innerHTML = content;
    // id
    a.setAttribute('id', idContent);
    // url
    a.setAttribute("href", url);

    return span;
}

function getEditorButton(content, idContent, url) {
    let span = document.createElement("span");
    span.setAttribute("class", "viewerbtn full-width");

    let a = document.createElement("a");
    a.setAttribute("class", "btn btn-success btn-lg btn-block");
    a.setAttribute("target", "_blank");

    let i = document.createElement("i");
    i.setAttribute("class", "fa fa-edit");
    a.appendChild(i);
    span.appendChild(a);

    // text
    a.innerHTML = content;
    // id
    a.setAttribute('id', idContent);
    // url
    a.setAttribute("href", url);

    return span;
}

window.onValidateWork = function () {

    $('#workButton').addClass("hidden");    
    
    if ($('#collapseOnlineTable').hasClass('active')) {
        
        if ($("#jsonData").val().length == 0) {
            showAlert(window.Resources["ChartOnlineTableIsEmptyMessage"]);
            $('#workButton').removeClass("hidden");

            return false;               
        }

        $('.add-property').toggle(false);
        $('.saveas').toggle(false);

        return true; 
    }      

    if ($('input[name="FileNames"]:isEmpty').length >= 0 &&
        $('input[name="FileNames"]').length == 0) {
        showAlert(window.Resources["FileSelectMessage"]);
        $('#workButton').removeClass("hidden");

        return false;
    }

    return true;
}

window.registerWorkFormHandler = () => $('.workForm').ajaxForm({
    url: APIMethodWorkUrl,
    dataType: "json",
    beforeSubmit: onValidateWork,
    beforeSend: showLoader,
    success: workSuccess,
    complete: hideLoader,
    error: handleWorkError
});

window.registerErrorPublishingFormHandler = () => $('.errorForm').ajaxForm({
    url: `${APIBasePath}api/Error2Forum/ReportError/`,
    dataType: "json",
    beforeSend: showLoader,
    beforeSubmit: onValidateReportError,
    success: reportErrorSuccess,
    complete: hideLoader,
    error: handleReportError
});

window.registerFormHandlers = function () {
    registerWorkFormHandler();
    registerErrorPublishingFormHandler();
};