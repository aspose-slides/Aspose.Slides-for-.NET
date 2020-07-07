using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using Aspose.Slides.Live.Demos.UI.Helpers;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Splits presentation to parts and saves each part to the specified format.
		/// </summary>
		/// <param name="source">The source presentation file.</param>
		/// <param name="outputDir">Output directory where parts will be stored.</param>
		/// <param name="format">The required format.</param>
		/// <param name="splitType">Splitting type <see cref="SplitType"/></param>
		/// <param name="splitNumber">The number of slides in the group (applied only for <see cref="SplitType.Number"/>)</param>
		/// <param name="splitRange">The slide ranges string (applied only for <see cref="SplitType.Range"/>)</param>
		/// <param name="cancellationToken">The cancellation token.</param>
		public void Split(string source,
			string outputDir,
			SlidesConversionFormat format,
			SplitType splitType,
			int splitNumber,
			string splitRange,
			CancellationToken cancellationToken = default(CancellationToken))
		{
			using (var presentation = new Presentation(source))
			{
				foreach (var chunk in presentation.GetChunks(splitType, splitNumber, splitRange))
				{
					SaveChunk(outputDir, chunk.name, Path.GetFileNameWithoutExtension(source), chunk.slides, format, cancellationToken);
				}
			}
		}

		private void SaveChunk(string outputDir, string chunkName, string fileName, ISlide[] chunkSlides, SlidesConversionFormat format, CancellationToken cancellationToken = default(CancellationToken))
		{
			if (chunkSlides.Length == 0)
			{
				return;
			}
			cancellationToken.ThrowIfCancellationRequested();

			var presentation = chunkSlides[0].Presentation;

			switch (format)
			{
				case SlidesConversionFormat.pptx:
				case SlidesConversionFormat.pptm:
				case SlidesConversionFormat.ppsx:
				case SlidesConversionFormat.ppsm:
				case SlidesConversionFormat.potx:
				case SlidesConversionFormat.potm:
				case SlidesConversionFormat.pps:
				case SlidesConversionFormat.ppt:
				case SlidesConversionFormat.pot:
				case SlidesConversionFormat.odp:
				case SlidesConversionFormat.otp:

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

				case SlidesConversionFormat.pdf:
				case SlidesConversionFormat.xps:
				case SlidesConversionFormat.tiff:
				case SlidesConversionFormat.html:
				case SlidesConversionFormat.swf:

					presentation.Save(Path.Combine(outputDir, $"{fileName}_{chunkName}.{format}"),
						chunkSlides.Select(s => s.SlideNumber).ToArray(),
						format.ToString().ParseEnum<SaveFormat>());
					break;

				case SlidesConversionFormat.doc:
				case SlidesConversionFormat.docx:

					using (var stream = new MemoryStream())
					{
						presentation.Save(stream, chunkSlides.Select(s => s.SlideNumber).ToArray(), SaveFormat.Html);
						stream.Flush();
						stream.Seek(0, SeekOrigin.Begin);

						var doc = new Words.Document(stream);
						var wordFormat = format == SlidesConversionFormat.doc
							? Words.SaveFormat.Doc
							: Words.SaveFormat.Docx;
						doc.Save(Path.Combine(outputDir, $"{fileName}_{chunkName}.{format}"), wordFormat);
					}
					break;

				case SlidesConversionFormat.txt:

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

				case SlidesConversionFormat.bmp:
				case SlidesConversionFormat.jpeg:
				case SlidesConversionFormat.png:
				case SlidesConversionFormat.emf:
				case SlidesConversionFormat.wmf:
				case SlidesConversionFormat.gif:
				case SlidesConversionFormat.exif:
				case SlidesConversionFormat.ico:

					foreach (var slide in chunkSlides)
					{
						cancellationToken.ThrowIfCancellationRequested();

						var outFile = Path.Combine(outputDir, $"{fileName}_{slide.SlideNumber:D2}.{format}");
						using (var bitmap = slide.GetThumbnail(1, 1))
						{
							bitmap.Save(outFile, GetImageFormat(format));
						}
					}
					break;

				case SlidesConversionFormat.svg:

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
