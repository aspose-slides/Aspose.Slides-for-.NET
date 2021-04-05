using System;
using System.Collections.Generic;
using System.Linq;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.UI.Models.Interfaces;

namespace Aspose.Slides.Web.UI.Models
{
    class ChartModel : BaseViewModel, IChartModel
    {
        public Dictionary<ChartTypes, string> ChartTypes { get; } =
            Enum.GetValues<ChartTypes>().ToDictionary(ct => ct, ct => ct.GetDescription());
        public (ChartTypes, string) ChartType { get; } = (API.Clients.Enums.ChartTypes.ClusteredColumn, API.Clients.Enums.ChartTypes.ClusteredColumn.GetDescription());
        public string SaveFormat { get; } = SlidesConversionFormats.png.ToString();
        public IEnumerable<string> Formats { get; } = Enum.GetNames(typeof(SlidesConversionFormats)).ToList();
        public string OnlineTab { get; } = "Online Table";
        public string UploadTab { get; } = "Upload File";
        public string PreviewAltText { get; } = "Chart preview image";
        public string TemplateButtonText { get; } = "Download Template";
        public string PreviewButtonText { get; } = "Generate Preview";

        public string HelpStep1 { get; } =
            "1. Choose one of two: press \"Upload File\" to upload the table data file from your device, or \"Online Table\" - to create table data online.";

        public string HelpStep2Upload { get; } = "2. Download file with the table template, edit it with your data and then upload it at the file drop area below.";
        public string HelpStep2Online { get; } = "2. Fill table with chart data to generate the chart on the image below.";
        public string HelpStep3 { get; } = "3. Choose the chart type and press \"Generate Preview\" to generate it.";
        public string HelpStep4 { get; } = "4. Choose the generate chart type and the file format to save it. Then press \"Create Chart\".";
    }
}