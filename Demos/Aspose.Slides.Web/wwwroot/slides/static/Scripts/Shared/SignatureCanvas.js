function Canvas(signCanvas, getColor) {
    var empty = true;
    var signatureCanvas = document.querySelector(signCanvas);
    if (signatureCanvas === null)
        return;
    var ctx = signatureCanvas.getContext('2d');

    var mouse = { x: null, y: null };
    var last_mouse = { x: 0, y: 0 };

    this.resetCanvasContext = function () {
        ctx.lineWidth = 2;
        ctx.lineJoin = 'round';
        ctx.lineCap = 'round';
        ctx.strokeStyle = getColor();
    }

    this.isEmpty = () => empty;

    this.resetCanvasContext();

    // set the listener on pointer down event
    var setOnDownListener = function (handlerDown, handlerMove) {
        signatureCanvas.addEventListener(handlerDown,
            function (ev) {
                signatureCanvas.addEventListener(handlerMove, onPaint, false);
            },
            false);
    };

    // set the listener on pointer up event
    var setOnUpListener = function (handlerUp, handlerMove) {
        signatureCanvas.addEventListener(handlerUp,
            function () {
                mouse.x = null;
                mouse.y = null;
                signatureCanvas.removeEventListener(handlerMove, onPaint, false);
            },
            false);
    };

    // set the listener on pointer move event
    var setOnMoveListener = function (handlerMove) {
        signatureCanvas.addEventListener(handlerMove,
            function (ev) {
                last_mouse.x = mouse.x;
                last_mouse.y = mouse.y;

                var bRect = signatureCanvas.getBoundingClientRect();
                mouse.x = (ev.clientX - bRect.left) * (signatureCanvas.width / bRect.width);
                mouse.y = (ev.clientY - bRect.top) * (signatureCanvas.height / bRect.height);
            },
            false);
    };

    // check if the specified event handler is supported
    var isEventSupported = function (eventName) {
        var cnv = document.createElement('canvas');
        var isSupported = (eventName in cnv);
        if (!isSupported) {
            cnv.setAttribute(eventName, 'return;');
            isSupported = typeof cnv[eventName] === 'function';
        }
        cnv = null;
        return isSupported;
    };

    // handler names on pointer events
    var handlerDown = null, handlerUp = null, handlerMove = null;

    // add event listeners
    if (isEventSupported('onpointerdown')) {
        handlerDown = 'pointerdown';
        handlerUp = 'pointerup';
        handlerMove = 'pointermove';
    } else if (isEventSupported('ontouchstart')) {
        handlerDown = 'touchstart';
        handlerUp = 'touchend';
        handlerMove = 'touchmove';
    } else if (isEventSupported('onmousedown')) {
        handlerDown = 'mousedown';
        handlerUp = 'mouseup';
        handlerMove = 'mousemove';
    }
    if (handlerDown !== null && handlerUp !== null && handlerMove !== null) {
        setOnDownListener(handlerDown, handlerMove);
        setOnUpListener(handlerUp, handlerMove);
        setOnMoveListener(handlerMove);
    }

    var onPaint = function () {
        if (last_mouse.x === null && last_mouse.y === null) {
            last_mouse.x = mouse.x;
            last_mouse.y = mouse.y;
        }
        ctx.beginPath();
        ctx.moveTo(last_mouse.x, last_mouse.y);
        ctx.lineTo(mouse.x, mouse.y);
        ctx.closePath();
        ctx.stroke();
        empty = false;
    };

    this.clearSignature = function () {
        empty = true;
        signatureCanvas.width = signatureCanvas.width;
        ctx = signatureCanvas.getContext('2d');
        ctx.lineWidth = 2;
        ctx.lineJoin = 'round';
        ctx.lineCap = 'round';
        ctx.strokeStyle = getColor();
        mouse.x = null;
        mouse.y = null;
    };
}
