using Aspose.Slides.Export;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of splitte logic.
	/// </summary>
	internal sealed class SplitterService : SlidesServiceBase, ISplitterService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public SplitterService(ILogger<SplitterService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Splits presentation to parts and saves each part to the specified format.
		/// </summary>
		/// <param name="source">The source presentation file.</param>
		/// <param name="outputDir">Output directory where parts will be stored.</param>
		/// <param name="format">The required format.</param>
		/// <param name="splitType">Splitting type <see cref="SplitTypes"/></param>
		/// <param name="splitNumber">The number of slides in the group (applied only for <see cref="SplitTypes.Number"/>)</param>
		/// <param name="splitRange">The slide ranges string (applied only for <see cref="SplitTypes.Range"/>)</param>
		/// <param name="cancellationToken">The cancellation token.</param>
		public void Split(string source,
			string outputDir,
			SlidesConversionFormats format,
			SplitTypes splitType,
			int? splitNumber,
			string splitRange,
			CancellationToken cancellationToken = default)
		{
			using (var presentation = new Presentation(source))
			{
				foreach (var chunk in presentation.GetChunks(splitType, splitNumber, splitRange))
				{
					SaveChunk(outputDir, chunk.name, Path.GetFileNameWithoutExtension(source), chunk.slides, format, cancellationToken);
				}
			}
		}

		private void SaveChunk(string outputDir, string chunkName, string fileName, ISlide[] chunkSlides, SlidesConversionFormats format, CancellationToken cancellationToken = default)
		{
			if (chunkSlides.Length == 0)
			{
				return;
			}
			cancellationToken.ThrowIfCancellationRequested();

			var presentation = chunkSlides[0].Presentation;

			switch (format)
			{
				case SlidesConversionFormats.pptx:
				case SlidesConversionFormats.pptm:
				case SlidesConversionFormats.ppsx:
				case SlidesConversionFormats.ppsm:
				case SlidesConversionFormats.potx:
				case SlidesConversionFormats.potm:
				case SlidesConversionFormats.pps:
				case SlidesConversionFormats.ppt:
				case SlidesConversionFormats.pot:
				case SlidesConversionFormats.odp:
				case SlidesConversionFormats.otp:

					using (var splitPresentation = new Presentation())
					{
						while (splitPresentation.Slides.Count > 0)
						{
							splitPresentation.Slides.RemoveAt(0);
						}

						foreach (var slide in chunkSlides)
						{
							cancellationToken.ThrowIfCancellationRequested();

							splitPresentation.Slides.AddClone(slide);
						}

						splitPresentation.Save(Path.Combine(outputDir, $"{fileName}_{chunkName}.{format}"), format.ToString().ParseEnum<SaveFormat>());
					}
					break;

				case SlidesConversionFormats.pdf:
				case SlidesConversionFormats.xps:
				case SlidesConversionFormats.tiff:
				case SlidesConversionFormats.html:
				case SlidesConversionFormats.swf:

					presentation.Save(Path.Combine(outputDir, $"{fileName}_{chunkName}.{format}"),
						chunkSlides.Select(s => s.SlideNumber).ToArray(),
						format.ToString().ParseEnum<SaveFormat>());
					break;

				case SlidesConversionFormats.doc:
				case SlidesConversionFormats.docx:

					using (var stream = new MemoryStream())
					{
						presentation.Save(stream, chunkSlides.Select(s => s.SlideNumber).ToArray(), SaveFormat.Html);
						stream.Flush();
						stream.Seek(0, SeekOrigin.Begin);

						var doc = new Words.Document(stream);
						var wordFormat = format == SlidesConversionFormats.doc
							? Words.SaveFormat.Doc
							: Words.SaveFormat.Docx;
						doc.Save(Path.Combine(outputDir, $"{fileName}_{chunkName}.{format}"), wordFormat);
					}
					break;

				case SlidesConversionFormats.txt:

					var lines = new List<string>();
					foreach (var slide in chunkSlides)
					{
						cancellationToken.ThrowIfCancellationRequested();

						foreach (var shp in slide.Shapes)
						{
							if (shp is AutoShape ashp)
							{
								lines.Add(ashp.TextFrame.Text);
							}
						}

						var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

						if (!string.IsNullOrEmpty(notes))
						{
							lines.Add(notes);
						}
					}
					System.IO.File.WriteAllLines(Path.Combine(outputDir, $"{fileName}_{chunkName}.{format}"), lines);
					break;

				case SlidesConversionFormats.bmp:
				case SlidesConversionFormats.jpeg:
				case SlidesConversionFormats.png:
				case SlidesConversionFormats.emf:
				case SlidesConversionFormats.wmf:
				case SlidesConversionFormats.gif:
				case SlidesConversionFormats.exif:
				case SlidesConversionFormats.ico:

					foreach (var slide in chunkSlides)
					{
						cancellationToken.ThrowIfCancellationRequested();

						var outFile = Path.Combine(outputDir, $"{fileName}_{slide.SlideNumber:D2}.{format}");
						using (var bitmap = slide.GetThumbnail(1, 1))
						{
							bitmap.Save(outFile, format.GetImageFormat());
						}
					}
					break;

				case SlidesConversionFormats.svg:

					var svgOptions = new SVGOptions
					{
						PicturesCompression = PicturesCompression.DocumentResolution
					};
					foreach (var slide in chunkSlides)
					{
						cancellationToken.ThrowIfCancellationRequested();

						var outFile = Path.Combine(outputDir, $"{fileName}_{slide.SlideNumber:D2}.{format}");
						using (var stream = new FileStream(outFile, FileMode.CreateNew))
						{
							slide.WriteAsSvg(stream, svgOptions);
						}
					}
					break;

				default:
					throw new ArgumentException($"Unknown format {format}");
			}
		}
	}
}
