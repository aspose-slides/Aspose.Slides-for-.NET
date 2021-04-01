using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class ComparisonModel : BaseViewModel, IComparisonModel
    {
        public IUploadFileModel SecondUploadFile { get; } = new UploadFileModel();
        public IEnumerable<string> ComparisonMethods { get; } = Enum.GetNames(typeof(ComparisonMethods)).ToList();
        public string ComparisonMethod { get; } = Aspose.Slides.Web.API.Clients.Enums.ComparisonMethods.BySlides.ToString();
        public IEnumerable<string> SaveFormats { get; } = Enum.GetNames(typeof(ComparisonDiffFileSaveFormats)).ToList();
        public string SaveFormat { get; } = ComparisonDiffFileSaveFormats.Pdf.ToString();
        public string LeftSideText { get; } = "The left compare file";
        public string RightSideText { get; } = "The right compare file";
    }
}