using System.Collections.Generic;
using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models.Slides
{
	public class PresentationInfo
	{
		///<Summary>
		/// get or set guid
		///</Summary>
		[JsonProperty("guid")]
		public string Guid { get; set; }

		///<Summary>
		/// get or set pages
		///</Summary>
		[JsonProperty("pages")]
		public List<PresentationSlide> Slides { get; set; }

		///<Summary>
		/// get or set printAllowed
		///</Summary>
		[JsonProperty("printAllowed")]
		public bool PrintAllowed { get; set; } = true;

		///<Summary>
		/// List of items for Navigation Pane
		///</Summary>
		[JsonProperty("navigationPane")]
		public List<PresentationNavigationItem> NavigationItems;
	}
}
