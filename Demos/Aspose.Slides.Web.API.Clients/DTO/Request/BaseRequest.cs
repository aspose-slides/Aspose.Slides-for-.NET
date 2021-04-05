using System.Collections.Generic;

namespace Aspose.Slides.Web.API.Clients.DTO.Request
{
	/// <summary>
	/// Base request model.
	/// </summary>
	public class BaseRequest
	{
		/// <summary>
		/// Upload id.
		/// </summary>
		public string id { get; set; }

		/// <summary>
		/// File names for processing.
		/// </summary>
		public IList<string> FileNames { get; set; }
	}
}
