using Aspose.Slides.Web.API.Clients.Enums;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Model for Import Request
	/// </summary>
	public sealed class ImportRequest : BaseRequest
	{
		/// <summary>
		/// Format of presentation
		/// </summary>
		public PresentationFormats SaveFormat { get; set; }
	}
}
