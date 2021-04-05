namespace Aspose.Slides.Web.API.Clients.DTO.Request
{	
	/// <summary>
	/// ProtectModel request model.
	/// </summary>
	public sealed class ProtectOptionsRequest : BaseRequest
	{
		/// <summary>
		/// Password for view.
		/// </summary>
		public string PasswordView { get; set; }

		/// <summary>
		/// Password for edit.
		/// </summary>
		public string PasswordEdit { get; set; }

		/// <summary>
		/// Mark file as read-only.
		/// </summary>
		public bool MarkAsReadonly { get; set; }

		/// <summary>
		/// Mark file as final.
		/// </summary>
		public bool MarkAsFinal { get; set; }
	}
}
