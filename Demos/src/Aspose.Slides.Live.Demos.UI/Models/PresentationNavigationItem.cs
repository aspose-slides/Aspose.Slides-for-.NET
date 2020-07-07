using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	public class PresentationNavigationItem
	{
		/// <summary>
		/// Name of the heading
		/// </summary>
		[JsonProperty("name")]
		public string Name { get; set; }

		/// <summary>
		/// Style: Heading1, Heading2, etc.
		/// </summary>
		[JsonProperty("style")]
		public int Style { get; set; }

		/// <summary>
		/// Page, on which the heading is
		/// </summary>
		[JsonProperty("pageNumber")]
		public int Number { get; set; }
	}
}
