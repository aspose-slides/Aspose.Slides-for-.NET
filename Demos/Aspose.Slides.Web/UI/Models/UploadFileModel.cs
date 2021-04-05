using System;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class UploadFileModel : IUploadFileModel
    {
        public bool AcceptMultipleFiles { get; set; } = false;
        public string AcceptedExtensions { get; set; }
        public string UploadId { get; } = Guid.NewGuid().ToString();
        public string Label { get; set; } = "Drop or upload your file";
    }
}