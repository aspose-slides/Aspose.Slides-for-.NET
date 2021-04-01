using System;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class ConversionModel : BaseViewModel, IConversionModel
    {
        public string FromFormat { get; } = SlidesConversionFormats.pptx.ToString();
        public string ToFormat { get; } = SlidesConversionFormats.pdf.ToString();

        public string[] Formats { get; } = Enum.GetValues(typeof(SlidesConversionFormats)).Cast<Enum>()
            .Select(s => s.ToString()).ToArray();
    }
}