using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	public class ViewerSlideResult : BaseResult
	{
		[JsonProperty("slide")]
		public PresentationSlide Slide{ get; set; }
	}
}
