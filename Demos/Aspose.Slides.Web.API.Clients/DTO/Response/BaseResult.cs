namespace Aspose.Slides.Web.API.Clients.DTO.Response
{
	/// <summary>
	/// Base class for results.
	/// </summary>
	public class BaseResult
	{
		/// <summary>
		/// Is result success?
		/// </summary>
		public bool IsSuccess { get;  set; }

		/// <summary>
		/// idError.
		/// </summary>
		[System.Diagnostics.CodeAnalysis.SuppressMessage("Style", "IDE1006:Naming Styles", Justification = "id always small")]
		public string idError { get;  set; }

		/// <summary>
		/// Message
		/// </summary>
		public string Message { get; set; }
	}
}
