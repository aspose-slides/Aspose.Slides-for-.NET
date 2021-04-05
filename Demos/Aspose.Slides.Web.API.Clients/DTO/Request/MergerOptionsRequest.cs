namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Merger options.
	/// </summary>
	public class MergerOptionsRequest : BaseRequest
	{
		/// <summary>
		/// Upload id for StyleMaster.
		/// </summary>
		public string idStyleMaster { get; set; }

		/// <summary>
		/// File name for StyleMaster.
		/// </summary>
		public string FileNameStyleMaster { get; set; }

		/// <summary>
		/// Format to convert into.
		/// When not specified, used pptx.
		/// </summary>
		public string Format { get; set; }
	}
}
