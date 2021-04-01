(function ($) {
    $.FileUploader = function (fileInput, parent, onChanged) {
        this.files = [];
        this.fileInput = fileInput;
        this.parent = parent;
        $(parent).change(() => onChanged());

        this.moveUp = function (file) {
            const index = this.files.indexOf(file);
            if (index > 0) {
                this.changePosition(index, index - 1);
            }
        }

        this.moveDown = function (file) {
            const index = this.files.indexOf(file);
            if (index < this.files.length - 1) {
                this.changePosition(index, index + 1);
            }
        }

        this.changePosition = function (from, to) {
            const cutOut = this.files.splice(from, 1)[0];
            this.files.splice(to, 0, cutOut);
            this.display();
        }

        this.remove = function (file) {
            const index = this.files.indexOf(file);
            if (index > -1) {
                this.files.splice(index, 1);
                this.display();
            }
        }

        function preventFileDrop(evt) {
            evt = evt || event;
            evt.preventDefault();
            evt.stopPropagation();
        };

        this.display = function () {
            $(this.parent).find('.fileupload').remove();

            for (let i = 0; i < this.files.length; i++) {
                let file = this.files[i];

                var fileMoveUpLink = null;
                var fileMoveDownLink = null;
                if (this.isMultiple()) {
                    fileMoveUpLink = $('\
                    <a class="fileMoveUpLink">\
                        <i class="fa fa-arrow-up"></i>\
                    </a>\
                    ');
                    fileMoveDownLink = $('\
                    <a class="fileMoveDownLink">\
                        <i class="fa fa-arrow-down"></i>\
                    </a>\
                    ');
                    fileMoveUpLink.find('i').on('click', $.proxy(function () {
                        this.moveUp(file);
                    }, this));
                    fileMoveDownLink.find('i').on('click', $.proxy(function () {
                        this.moveDown(file);
                    }, this));
                }
                var fileRemoveLink = $('\
                <a class="fileRemoveLink">\
                    <i class="fa fa-times"></i>\
                </a>\
                ');
                fileRemoveLink.find('i').on('click', $.proxy(function () {
                    this.remove(file);
                }, this));
                var spanFileName = $('\
                <span class="filename">\
                    <label class="custom-file-upload" style="display:inline">' + file + '</label>\
                </span>\
                ');
                if (fileMoveUpLink !== null && fileMoveDownLink !== null) {
                    spanFileName.append(fileMoveUpLink);
                    spanFileName.append(fileMoveDownLink);
                }
                spanFileName.append(fileRemoveLink);

                var fileBlock = $('<div class="fileupload"></div>');
                fileBlock.on('dragover', preventFileDrop);
                fileBlock.on('drop', preventFileDrop);
                fileBlock.append(spanFileName);

                this.parent.append(fileBlock);
            }
        }

        this.isMultiple = () => Boolean($(this.fileInput).attr("multiple"));

        this.appendFiles = function (uploaded) {
            this.files = this.files.concat(uploaded);
            this.display();
        }

        this.replaceFiles = function (uploaded) {
            this.files = uploaded;
            this.display();
        }
    };
})(jQuery);
