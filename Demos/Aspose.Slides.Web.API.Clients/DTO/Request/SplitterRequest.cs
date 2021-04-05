using System.Text.Json.Serialization;
using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Splitter request model.
	/// </summary>
	public class SplitterRequest : BaseRequest
	{
		/// <summary>
		/// Split type <see cref="SplitType"/>.
		/// </summary>
		[JsonPropertyName("splitType")]
		public SplitTypes SplitType { get; set; }

		/// <summary>
		/// Split ranges string (applicable only for <see cref="SplitType.Range"/>).
		/// </summary>
		[JsonPropertyName("splitRange")]
		public string SplitRange { get; set; }

		/// <summary>
		/// Split group size (applicable only for <see cref="SplitType.Number"/>).
		/// </summary>
		[JsonPropertyName("splitNumber")]
		public int? SplitNumber { get; set; }

		/// <summary>
		/// The partition file format.
		/// </summary>
		[JsonPropertyName("format")]
		public string Format { get; set; }
	}
}
