using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	public interface ITemporaryFolder : IDisposable
	{
		string FolderName { get; }
		Task<string> SaveAsync(Stream stream, string fileName, CancellationToken cancellationToken = default);
		MemoryStream GetArchiveStream();
	}
}
