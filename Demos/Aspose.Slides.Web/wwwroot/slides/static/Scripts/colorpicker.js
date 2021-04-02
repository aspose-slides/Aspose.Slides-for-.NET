var colorList = [ '000000', '993300', '333300', '003300', '003366', '000066', '333399', '333333', 
'660000', 'FF6633', '666633', '336633', '336666', '0066FF', '666699', '666666', 'CC3333', 'FF9933', '99CC33', '669966', '66CCCC', '3366FF', '663366', '999999', 'CC66FF', 'FFCC33', 'FFFF66', '99FF66', '99CCCC', '66CCFF', '993366', 'CCCCCC', 'FF99CC', 'FFCC99', 'FFFF99', 'CCffCC', 'CCFFff', '99CCFF', 'CC99FF', 'FFFFFF' ];

function colorPicker(pickerSelector = '#color-picker') {
    var picker = $(pickerSelector);

    for (var i = 0; i < colorList.length; i++) {
        picker.append('<li class="color-item" data-hex="' +
            '#' +
            colorList[i] +
            '" style="background-color:' +
            '#' +
            colorList[i] +
            ';"></li>');
    }

    const color = picker.parent().find('.color-picker-text').val();
    picker.parent().find('.color-holder').css('background-color', color);

    $('body').click(function() {
        picker.fadeOut();
    });

    picker.parent().find('.call-picker').click(function(event) {
        event.stopPropagation();
        picker.fadeIn();
        picker.children('li').hover(function() {
            var codeHex = $(this).data('hex');

            picker.parent().find('.color-holder').css('background-color', codeHex);
            picker.parent().find('.color-picker-text').val(codeHex);
            picker.parent().find('.color-picker-text').change();
        });
    });
}