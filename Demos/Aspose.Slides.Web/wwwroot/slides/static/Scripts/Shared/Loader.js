window.showLoader = function () {
    $('#loader').removeClass("hidden");
    hideAlert();
};

window.hideLoader = function () {
    $('#loader').addClass("hidden");
};