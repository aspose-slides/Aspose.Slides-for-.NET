using System.Text.Json.Serialization;

namespace Aspose.Slides.Web.API.Clients.DTO.Response
{
	/// <summary>
	/// Viewer info request result.
	/// </summary>
	public class ViewerInfoResult : BaseResult
	{
		/// <summary>
		/// Presentation information.
		/// </summary>
		[JsonPropertyName("info")]
		public PresentationInfoDTO Info { get; set; }
	}
}
