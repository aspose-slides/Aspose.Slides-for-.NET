using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class ImportModel : BaseViewModel, IImportModel
    {
        public string SaveFormat { get; } = PresentationFormats.Pptx.ToString();
        public IEnumerable<string> SaveFormats { get; } = Enum.GetNames(typeof(PresentationFormats)).ToList();
    }
}