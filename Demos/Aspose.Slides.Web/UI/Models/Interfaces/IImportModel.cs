using System.Collections.Generic;

namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IImportModel : IBaseViewModel
	{
		string SaveFormat { get; }
		IEnumerable<string> SaveFormats { get; }
	}
}
