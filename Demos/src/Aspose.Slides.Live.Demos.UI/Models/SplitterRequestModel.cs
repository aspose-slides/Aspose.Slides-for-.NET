using Newtonsoft.Json;

namespace Aspose.Slides.Live.Demos.UI.Models
{
	/// <summary>
	/// Splitter request model.
	/// </summary>
	public class SplitterRequestModel : BaseRequestModel
	{
		/// <summary>
		/// Split type <see cref="SplitType"/>.
		/// </summary>
		[JsonProperty("splitType")]
		public SplitType SplitType { get; set; }

		/// <summary>
		/// Split ranges string (applicable only for <see cref="SplitType.Range"/>).
		/// </summary>
		[JsonProperty("splitRange")]
		public string SplitRange { get; set; }

		/// <summary>
		/// Split group size (applicable only for <see cref="SplitType.Number"/>).
		/// </summary>
		[JsonProperty("splitNumber")]
		public int SplitNumber { get; set; }

		/// <summary>
		/// The partition file format.
		/// </summary>
		[JsonProperty("format")]
		public string Format { get; set; }

	}
}
