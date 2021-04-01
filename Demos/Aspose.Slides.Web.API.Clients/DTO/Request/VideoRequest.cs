using System.Text.Json.Serialization;
using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Video request model.
	/// </summary>
	public class VideoRequest : BaseRequest
	{
		/// <summary>
		/// Video codec <see cref="VideoCodecs"/>.
		/// </summary>
		[JsonPropertyName("videoCodec")]
		public VideoCodecs VideoCodec { get; set; }

		/// <summary>
		/// Split ranges string
		/// </summary>
		[JsonPropertyName("splitRange")]
		public string SplitRange { get; set; }

		/// <summary>
		/// Transition time in seconds.
		/// </summary>
		[JsonPropertyName("transitionTime")]
		public int TransitionTime { get; set; }

	}
}
