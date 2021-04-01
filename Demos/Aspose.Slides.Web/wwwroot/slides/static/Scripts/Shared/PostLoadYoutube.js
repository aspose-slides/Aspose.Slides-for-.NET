function init() {
    var vidDefer = document.getElementsByTagName('iframe');
    for (var i = 0; i < vidDefer.length; i++) {
        if (vidDefer[i].getAttribute('post-load-src')) {
            vidDefer[i].setAttribute('src', vidDefer[i].getAttribute('post-load-src'));
        }
    }
}
window.onload = init;