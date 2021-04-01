using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Slides.Web.Interfaces.Services;

namespace Aspose.Slides.Web.Core.Services.Storage
{
	internal class SourceFileStorage : ISourceStorage
	{
		private readonly LocalFileStorage _fileStorage;

		public SourceFileStorage(string root)
		{
			_fileStorage = new LocalFileStorage(root);
		}

		public Task<Stream> GetStreamAsync(string folder, string filename, CancellationToken cancellationToken = default)
		{
			return _fileStorage.GetStreamAsync(folder, filename, cancellationToken);
		}

		public Task UploadAsync(string folder, string filename, Stream stream, CancellationToken cancellationToken = default)
		{
			return _fileStorage.UploadAsync(folder, filename, stream, cancellationToken);
		}

		public Task<IEnumerable<string>> ListFilesAsync(string folder, CancellationToken cancellationToken)
		{
			return _fileStorage.ListFilesAsync(folder, cancellationToken);
		}

		public Task<bool> IsExistAsync(string folder, string filename, CancellationToken cancellationToken = default)
		{
			return _fileStorage.IsExistAsync(folder, filename, cancellationToken);
		}
	}
}
