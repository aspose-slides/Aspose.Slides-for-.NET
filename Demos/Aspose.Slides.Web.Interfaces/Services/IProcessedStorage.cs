using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface to access processed files.
	/// </summary>
	public interface IProcessedStorage
	{
		Task<Stream> GetStreamAsync(string folder, string filename, CancellationToken cancellationToken = default);

		Task UploadAsync(string folder, string filename, Stream stream, CancellationToken cancellationToken= default);

		Task<bool> IsExistAsync(string folder, string filename, CancellationToken cancellationToken = default);
	}
}
