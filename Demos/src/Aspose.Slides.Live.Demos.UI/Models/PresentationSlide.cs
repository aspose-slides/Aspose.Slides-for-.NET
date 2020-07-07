using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	public class PresentationSlide
	{
		///<Summary>
		/// get or set width
		///</Summary>
		[JsonProperty("width")]
		public double Width { get; set; }

		///<Summary>
		/// get or set height
		///</Summary>
		[JsonProperty("height")]
		public double Height { get; set; }

		///<Summary>
		/// get or set number
		///</Summary>
		[JsonProperty("number")]
		public int Number { get; set; }

		///<Summary>
		/// get or set angle
		///</Summary>
		[JsonProperty("angle")]
		public int Angle { get; set; }

		///<Summary>
		/// get or set data
		///</Summary>
		[JsonProperty("data")]
		public string Data { get; set; }
	}
}
