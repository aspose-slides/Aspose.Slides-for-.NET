using Aspose.Slides.Web.Interfaces.Services;
using Aspose.Slides.Web.Core.Helpers;
using System;
using System.IO;
using System.IO.Compression;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services.Storage
{
	internal class TemporaryFolder : ITemporaryFolder
	{
		private const int DEFAULT_BUFFER_SIZE = 81920;
		private readonly string _path;
		private readonly bool isDisposed = false;

		public TemporaryFolder(string path)
		{
			_path = path;
			FolderName = Path.GetFileName(path);
			Directory.CreateDirectory(path);
		}

		public string FolderName { get; }

		public void Dispose()
		{
			Dispose(true);
			GC.SuppressFinalize(this);
		}

		public async Task<string> SaveAsync(Stream stream, string fileName, CancellationToken token = default)
		{
			var filePath = Path.Combine(_path, fileName);
			using var fileStream = File.Create(filePath);
			await stream.CopyToAsync(fileStream, DEFAULT_BUFFER_SIZE, token);
			return filePath;
		}

		public MemoryStream GetArchiveStream()
		{
			
			var memoryStream = new MemoryStream();
			try
			{
				using var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true);
				archive.CreateEntryFromDirectory(_path);
				return memoryStream;
			}
			catch
			{
				memoryStream.Dispose();
				throw;
			}
		}

		public override string ToString()
		{
			return _path;
		}

		protected virtual void Dispose(bool disposing)
		{
			if (isDisposed)
			{
				return;
			}

			Directory.Delete(_path, true);
		}

		~TemporaryFolder()
		{
			Dispose(false);
		}
	}
}
