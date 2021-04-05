using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services.Storage
{
	internal class LocalFileStorage 
	{
		public string Root { get; }

		public LocalFileStorage(string root)
		{
			Root = root;
			if (!Directory.Exists(Root))
			{
				Directory.CreateDirectory(Root);
			}
		}

		public Task<Stream> GetStreamAsync(string folder, string filename, CancellationToken cancellationToken = default)
		{
			var path = Path.Combine(Root, folder, filename);
			if (!File.Exists(path))
			{
				return Task.FromResult<Stream>(null);
			}

			return Task.FromResult<Stream>(new FileStream(path, FileMode.Open, FileAccess.Read));
		}

		public async Task UploadAsync(string folder, string filename, Stream stream, CancellationToken cancellationToken = default)
		{
			var path = Path.Combine(Root, folder, filename);
			if (!Directory.Exists(Path.Combine(Root, folder)))
			{
				Directory.CreateDirectory(Path.Combine(Root, folder));
			}

			await using var fileStream = new FileStream(path, FileMode.Create, FileAccess.Write);
			stream.Seek(0, SeekOrigin.Begin);
			await stream.CopyToAsync(fileStream, cancellationToken);
		}

		public Task<IEnumerable<string>> ListFilesAsync(string folder, CancellationToken cancellationToken)
		{
			if (!Directory.Exists(Path.Combine(Root, folder)))
			{
				return Task.FromResult<IEnumerable<string>>(new string[]{});
			}
			return Task.FromResult<IEnumerable<string>>(Directory.GetFiles(Path.Combine(Root, folder)));
		}

		public Task<bool> IsExistAsync(string folder, string filename, CancellationToken cancellationToken = default)
		{
			var path = Path.Combine(Root, folder, filename);
			return Task.FromResult(File.Exists(path));
		}
	}
}
