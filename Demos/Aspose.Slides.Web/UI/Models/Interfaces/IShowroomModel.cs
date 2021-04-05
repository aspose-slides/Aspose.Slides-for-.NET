using System.Collections.Generic;

namespace Aspose.Slides.Web.UI.Models.Interfaces
{
	public interface IShowroomModel
	{
		List<Showcase> Showcases { get; set; }
		string IndexPageTitle { get; }
		string IndexPageSubTitle { get; }
		string ProductFamilyInclude { get; }
		string AsposeSlides { get; }
	}
}
