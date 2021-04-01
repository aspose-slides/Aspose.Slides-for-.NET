function removeFile(event) {
    let parent = $(event.target).closest('.filedrop');
    let fileNames = $('input[name="FileNames"]').toArray();
    let removableFile = $(event.target).prev('.custom-file-upload')[0];

    fileNames.forEach(upFile => {
        if (removableFile != null && $(upFile).val() === removableFile.innerText) {
            upFile.remove();
        }
    });

    $(event.target).closest('.fileupload').remove();
    parent.find('.uploadfileinput').removeClass("hidden");
    parent.find('.uploadfileinput').val("");

    afterRemoveFile(event);
}

function uploadFileSelected(event) {
    if (event.target.files[0] === undefined)
        return;

    hideAlert();

    var parent = $(event.target).closest('.filedrop');

    for (let i = 0; i < event.target.files.length; i++) {
        let file = event.target.files[i];
        parent.append(
            `<div class='fileupload'>
    <span class="filename">
        <a class="fileRemoveLink">
            <label for="UploadFileInput" class="custom-file-upload">${file.name}</label>
            <i class="fa fa-times"></i>
        </a>
    </span>
</div>`
        );
    }

    registerUploadHandlers();

    $(event.target).closest('.uploadForm.activeForm').submit();
};

function validateFilesSelection(arr, $form, options) {
    if ($form.find('.uploadfileinput:isEmpty').length > 0) {
        showAlert(window.Resources["FileSelectMessage"]);
        return false;
    }
}

function beforeUpload() {
    $('#workButton').addClass("hidden");
    showLoader();
}

function completeUpload(data, textStatus, xhr) {    
    $('#workButton').removeClass("hidden");
}

function uploadProgress(event, position, total, percentComplete) {
}

function uploadSuccess(data, textStatus, xhr) {
    hideLoader();
    const uploaded = data.filter(f => f.IsSuccess).map(f => f);   

    if (uploaded.length !== 0) {
        $('input[name="id"]').val(uploaded[0].id);
        $('#hdErrorFolder').val(uploaded[0].id);
        SEONavigate('uploaded');
    }

    data.forEach(upFile => {
        if (upFile.IsSuccess) {
            var fileName = $('<input type="hidden" name="FileNames"/>').val(upFile.FileName);
            $('input[name="id"][id!="hdErrorFolder"]').after(fileName);
        }
        else {
            dataErrorAlert(upFile);
            afterDataErrorAlert(upFile);
        }
    });

    if (uploaded.length !== 0) {
        afterUploadSuccess(uploaded, textStatus, xhr);
    }
}

window.registerUploadHandlers = function () {
    $('.uploadfileinput').unbind();
    $('.fileRemoveLink').unbind();
    $('.uploadfileinput').unbind();

    $('.uploadfileinput').click(removeFile);
    $('.fileRemoveLink').click(removeFile);
    $('.uploadfileinput').change(uploadFileSelected);

    $('.uploadForm.activeForm').ajaxForm({
        url: `${APIBasePath}api/Common/UploadFiles`,
        dataType: "json",
        beforeSubmit: validateFilesSelection,
        beforeSend: beforeUpload,
        uploadProgress: uploadProgress,
        success: uploadSuccess,
        complete: completeUpload,
        error: handleError
    });
};

window.afterUploadSuccess = function (data, textStatus, xhr) {
};

window.afterRemoveFile = function (event) {
};

window.afterDataErrorAlert = function (data) { };
