using Microsoft.AspNetCore.Http;

namespace Aspose.Slides.Web.API.Models
{
	public class UploadRequest
	{
		public IFormFileCollection UploadFileInput { get; set; }

		public string idUpload { get; set; }
	}
}
