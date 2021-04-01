using Aspose.Slides.Web.API.Clients.DTO.Response;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.UI.Services
{
	public interface IEditorService
	{
		Task<NewPresentationResponse> CreateByTemplateAsync(string template);
		Task<NewPresentationResponse> CopyProcessedAsync(string folder, string fileName);
	}
}
