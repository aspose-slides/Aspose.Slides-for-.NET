using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for uploaded source (unprocessed) files.
	/// </summary>
	public interface ISourceStorage
	{
		Task<Stream> GetStreamAsync(string folder, string filename, CancellationToken cancellationToken);

		Task UploadAsync(string folder, string filename, Stream stream, CancellationToken cancellationToken);

		Task<IEnumerable<string>> ListFilesAsync(string folder, CancellationToken cancellationToken);

		Task<bool> IsExistAsync(string folder, string filename, CancellationToken cancellationToken);
	}
}
