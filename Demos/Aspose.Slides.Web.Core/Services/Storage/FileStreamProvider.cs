using Aspose.Slides.Web.Interfaces.Services;
using System.IO;

namespace Aspose.Slides.Web.Core.Services.Storage
{
	internal class FileStreamProvider : IFileStreamProvider
	{
		public FileStreamProvider()
		{
		}

		public Stream GetStream(string path)
		{
			return File.OpenRead(path);
		}
	}
}
