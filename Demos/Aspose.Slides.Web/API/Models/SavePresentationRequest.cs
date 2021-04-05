using Microsoft.AspNetCore.Http;
using System.Text.Json.Serialization;

namespace Aspose.Slides.Web.API.Models
{
	public sealed class SavePresentationRequest
	{
		[JsonPropertyName("slidesData")]
		public IFormFileCollection SlidesData { get; set; }

		[JsonPropertyName("fileName")]
		public string FileName { get; set; }

		[JsonPropertyName("idUpload")]
		public string IdUpload { get; set; }

		[JsonPropertyName("slides")]
		public string Slides { get; set; }
	}
}
