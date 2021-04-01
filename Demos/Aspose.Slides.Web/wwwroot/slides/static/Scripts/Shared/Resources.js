window.Resources = [];

function getResourcesSuccess(data, textStatus, xhr) {
    Resources = data;
}

$(document).ready(function () {
    $.ajax({
        method: "POST",
        url: '/common/resources',
        dataType: "json",
        success: getResourcesSuccess,
        error: handleError
    });
});