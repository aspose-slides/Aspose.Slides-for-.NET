using System.Collections.Generic;
using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IChartModel : IBaseViewModel
	{
		Dictionary<ChartTypes, string> ChartTypes { get; }
		(ChartTypes, string) ChartType { get; }
		string SaveFormat { get; }
		IEnumerable<string> Formats { get; }
		string OnlineTab { get; }
		string UploadTab { get; }
		string PreviewAltText { get; }
		string TemplateButtonText { get; }
		string PreviewButtonText { get; }
		string HelpStep1 { get; }
		string HelpStep2Upload { get; }
		string HelpStep2Online { get; }
		string HelpStep3 { get; }
		string HelpStep4 { get; }
	}
}
