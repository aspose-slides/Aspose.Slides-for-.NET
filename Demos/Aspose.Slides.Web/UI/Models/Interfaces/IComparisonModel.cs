using System.Collections.Generic;

namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IComparisonModel : IBaseViewModel
	{
		IUploadFileModel SecondUploadFile { get; }
		IEnumerable<string> ComparisonMethods { get; }
		string ComparisonMethod { get; }
		IEnumerable<string> SaveFormats { get; }
		string SaveFormat { get; }
		string LeftSideText { get; }
		string RightSideText { get; }
	}
}
