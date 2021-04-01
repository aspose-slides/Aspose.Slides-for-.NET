window.SEONavigate = function (newPart) {
    var url = window.location.href;

    if (!url.includes('navigation/')) {
        if (!url.endsWith('/'))
            url = url + '/';
        url += 'navigation/';
    }

    if (!url.endsWith('/'))
        url = url + '/';

    url += newPart;

    window.history.pushState('pageNavigation', document.title, url);
};