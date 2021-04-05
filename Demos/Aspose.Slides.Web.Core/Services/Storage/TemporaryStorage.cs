using Aspose.Slides.Web.Interfaces.Services;
using System;
using System.IO;
using Microsoft.Extensions.Logging;

namespace Aspose.Slides.Web.Core.Services.Storage
{
	public class TemporaryStorage : SlidesServiceBase, ITemporaryStorage
	{
		public string Root { get; }

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="root"></param>
		public TemporaryStorage(ILogger<TemporaryStorage> logger, string root) : base(logger)
		{
			Root = root;
		}

		public ITemporaryFolder GetTemporaryFolder()
		{
			return GetTemporaryFolder(Guid.NewGuid().ToString());
		}

		public ITemporaryFolder GetTemporaryFolder(string id)
		{
			var info = Directory.CreateDirectory(Path.Combine(Root, id));
			return new TemporaryFolder(info.FullName);
		}
	}
}
