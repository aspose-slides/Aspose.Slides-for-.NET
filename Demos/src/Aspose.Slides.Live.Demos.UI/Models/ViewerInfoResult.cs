using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models.Slides
{
	public class ViewerInfoResult : BaseResult
	{
		[JsonProperty("info")]
		public PresentationInfo Info { get; set; }
	}
}
