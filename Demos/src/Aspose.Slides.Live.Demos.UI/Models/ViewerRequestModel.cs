using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// Viewer info request model
	/// </summary>
	public class ViewerRequestModel : BaseRequestModel
	{
		/// <summary>
		/// Gets and sets slide number (starts from 0)
		/// </summary>
		[JsonProperty("pageNumber")]
		public int SlideNumber { get; set; }

		public string FolderName { get; set; }

	}
}
