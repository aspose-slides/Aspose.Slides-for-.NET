using System;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class SplitterModel : BaseViewModel, ISplitterModel
    {
        public string IntoMany { get; } = "Split into many";
        public string Every { get; } = "Every slide";
        public string Odd { get; } = "Odd and even slides";
        public string ByNumber { get; } = "By slide number";
        public string IntoSingle { get; } = "Split into single";
        public string PageRange { get; } = "By slide range";
        public string ToFormat { get; } = SlidesConversionFormats.pdf.ToString();
        public string[] Formats { get; } = Enum.GetValues(typeof(SlidesConversionFormats)).Cast<Enum>().Select(s => s.ToString()).ToArray();
        public string RangeException { get; } = "Please, enter valid slide range in the field";
        public string NumberException { get; } = "Please, enter number in the field";
    }
}