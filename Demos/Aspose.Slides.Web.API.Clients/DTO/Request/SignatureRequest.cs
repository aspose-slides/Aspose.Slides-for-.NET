namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// The signature request model class.
	/// </summary>
	public class SignatureRequest : BaseRequest
	{
		/// <summary>
		/// Output format.
		/// </summary>
		public string Format { get; set; }

		/// <summary>
		/// Signature drawing base64 encoded.
		/// </summary>
		public string Drawing { get; set; }

		/// <summary>
		/// Signature text.
		/// </summary>
		public string Text { get; set; }

		/// <summary>
		/// HTML-like color string.
		/// </summary>
		public string Color { get; set; }

		/// <summary>
		/// Uploaded signature image folder name.
		/// </summary>
		public string idSignatureImage { get; set; }

		/// <summary>
		/// Uploaded signature image filename.
		/// </summary>
		public string FileNameSignatureImage { get; set; }
	}
}
