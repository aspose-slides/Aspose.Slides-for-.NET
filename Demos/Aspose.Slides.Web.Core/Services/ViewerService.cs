using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading;
using Viewer = Aspose.Slides.Web.Interfaces.Models.Viewer;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of viewer logic.
	/// </summary>
	internal sealed class ViewerService : SlidesServiceBase, IViewerService 
	{
		public string MarkerId => ".done";

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public ViewerService(ILogger<ViewerService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Copies source presentation to the output folder, generates SVG representation of slides and returns presentation information.
		/// </summary>
		/// <param name="id">The upload identifier.</param>
		/// <param name="sourceFile">The source presentation file path.</param>
		/// <param name="destinationPath">The resulting presentation file path.</param>
		/// <param name="cancellationToken">The cancellation token.</param>
		/// <returns>The slides information.</returns>
		public Viewer.PresentationInfo GetViewerInfo(string id, string sourceFile, string destinationPath, CancellationToken cancellationToken = default)
		{
			var outDir = Path.GetDirectoryName(destinationPath);
			var markerFile = $"{Path.GetFileName(sourceFile)}{MarkerId}";
			if (!File.Exists(Path.Combine(outDir, markerFile)))
			{
				using (Mutex mutex = new Mutex(false, id + Path.GetFileName(sourceFile)))
				{
					var mutexAcquired = false;
					try
					{
						try
						{
							mutexAcquired = mutex.WaitOne(Timeout.Infinite);
						}
						catch (AbandonedMutexException)
						{
							mutexAcquired = true;
						}

						cancellationToken.ThrowIfCancellationRequested();
						if (mutexAcquired)
						{
							// !!! we call sync File.Copy here because ReleaseMutex must be called from the same thread !!!
							if (File.Exists(sourceFile) && !File.Exists(destinationPath))
							{
								File.Copy(sourceFile, destinationPath, true);
							}
							return GenerateFiles(destinationPath, cancellationToken);
						}
					}
					finally
					{
						if (mutexAcquired)
						{
							mutex.ReleaseMutex();
						}
					}
				}
			}

			cancellationToken.ThrowIfCancellationRequested();
			using (var presentation = new Presentation(destinationPath))
			{
				return GetInfo(presentation);
			}
		}

		private Viewer.PresentationInfo GenerateFiles(string destination, CancellationToken cancellationToken = default)
		{
			var outDir = Path.GetDirectoryName(destination);
			Viewer.PresentationInfo result = null;
			var markerFile = $"{Path.GetFileName(destination)}{MarkerId}";

			if (File.Exists(Path.Combine(outDir, markerFile)))
			{
				return result;
			}

			cancellationToken.ThrowIfCancellationRequested();
			using (var presentation = new Presentation(destination))
			{
				cancellationToken.ThrowIfCancellationRequested();
				foreach (var slide in presentation.Slides)
				{
					GenerateSlide(Path.GetFileName(destination), slide, outDir, cancellationToken);

					cancellationToken.ThrowIfCancellationRequested();
				}

				result = GetInfo(presentation);
			}

			File.WriteAllText(Path.Combine(outDir, markerFile), "done");
			return result;
		}

		private static Viewer.PresentationInfo GetInfo(Presentation presentation)
		{
			return new Viewer.PresentationInfo
			{
				Width = (int)Math.Round(presentation.SlideSize.Size.Width),
				Height = (int)Math.Round(presentation.SlideSize.Size.Height),
				Count = presentation.Slides.Count
			};
		}

		private static void GenerateSlide(string fileName, ISlide slide, string outDir, CancellationToken cancellationToken = default)
		{
			cancellationToken.ThrowIfCancellationRequested();
			var slideFile = $"{fileName}.slide_{slide.SlideNumber}.svg";
			using (var stream = new MemoryStream())
			{
				slide.WriteAsSvg(stream);
				// !!! Workaround !!!
				// All generated svg have the ids starting with "page0",
				// we will replace them to "page{slideNumber}"
				stream.Flush();
				stream.Seek(0, SeekOrigin.Begin);

				using (var file = File.Create(Path.Combine(outDir, slideFile)))
				using (var writer = new StreamWriter(file))
				using (var reader = new StreamReader(stream))
				{
					while (!reader.EndOfStream)
					{
						var line = reader.ReadLine();
						if (slide.SlideNumber != 1)
						{
							line = line?.Replace("page0", $"page{slide.SlideNumber - 1}");
						}

						writer.WriteLine(line);
					}
				}
			}
		}
	}
}
