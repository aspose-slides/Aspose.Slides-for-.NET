using System.Text.Json.Serialization;

namespace Aspose.Slides.Web.API.Clients.DTO
{
	/// <summary>
	/// Presentation information class.
	/// </summary>
	public sealed class PresentationInfoDTO
	{
		/// <summary>
		/// Presentation width in pixels.
		/// </summary>
		[JsonPropertyName("width")]
		public int Width { get; set; }

		/// <summary>
		/// Presentation height in pixels.
		/// </summary>
		[JsonPropertyName("height")]
		public int Height { get; set; }

		/// <summary>
		/// Number of slides in the presentation.
		/// </summary>
		[JsonPropertyName("count")]
		public int Count { get; set; }
	}
}
