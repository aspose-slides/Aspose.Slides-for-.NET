using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	public class ViewerThumbnailsResult : BaseResult
	{
		[JsonProperty("thumbnails")]
		public PresentationSlide[] Thumbnails { get; set; }
	}
}
