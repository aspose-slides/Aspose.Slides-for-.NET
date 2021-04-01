using System;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class MergerModel : BaseViewModel, IMergerViewModel
    {
        public string StyleMasterUploadText { get; } = "Optionally upload style master file";
        public IUploadFileModel StyleMasterUploadFile { get; } = new UploadFileModel { Label = "Drop or upload your style master file" };
        public string ToFormat { get; } = SlidesConversionFormats.pdf.ToString();

        public string[] Formats { get; } = Enum.GetValues(typeof(SlidesConversionFormats)).Cast<Enum>()
            .Select(s => s.ToString()).ToArray();
    }
}