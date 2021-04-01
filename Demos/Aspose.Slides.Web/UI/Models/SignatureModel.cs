using System;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class SignatureModel : BaseViewModel, ISignatureModel
    {
        public string SelectFileText { get; } = "Please select a file to upload";
        public string SignatureDrawingError { get; } = "Draw the signature";
        public string SignatureTextError { get; } = "Enter the text";
        public string SignatureTypeDrawing { get; } = "Drawing";
        public string SignatureTypeText { get; } = "Text";
        public string SignatureTypeImage { get; } = "Image";
        public string AddSignatureText { get; } = "Add Your Text Signature";
        public string ToFormat { get; } = SlidesConversionFormats.pptx.ToString();
        public string[] Formats { get; } = Enum.GetValues(typeof(SlidesConversionFormats)).Cast<Enum>().Select(s => s.ToString()).ToArray();
        public IUploadFileModel UploadSignatureImage { get; } = new UploadFileModel { AcceptedExtensions = ".jpg|.jpeg|.png|.bmp" };
    }
}