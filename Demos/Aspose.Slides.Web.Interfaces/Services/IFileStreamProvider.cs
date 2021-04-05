using System.IO;

namespace Aspose.Slides.Web.Interfaces.Services
{
	public interface IFileStreamProvider
	{
		Stream GetStream(string path);
	}
}
