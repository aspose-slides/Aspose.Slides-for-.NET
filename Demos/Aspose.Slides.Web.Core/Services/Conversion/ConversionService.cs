using Aspose.Slides.Export;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;


namespace Aspose.Slides.Web.Core.Services.Conversion
{
	/// <summary>
	/// Implementation of slides conversion logic.
	/// </summary>
	internal sealed class ConversionService : SlidesServiceBase, IConversionService
	{
		private readonly IGifEncoder _gifEncoder;

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="gifEncoder"></param>
		/// <param name="licenseProvider"></param>
		public ConversionService(ILogger<ConversionService> logger,
			IGifEncoder gifEncoder,
			ILicenseProvider licenseProvider) : base(logger)
		{
			_gifEncoder = gifEncoder;

			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
			licenseProvider.SetAsposeLicense(AsposeProducts.Words);
		}

		/// <summary>
		/// Converts source file into target format, saves resulted file to out file.
		/// Returns null in case of multiple files (all saved into outFolder).
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>		
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file paths.</returns>
		public IEnumerable<string> Conversion(
			IList<string> sourceFiles,
			string outFolder,
			SlidesConversionFormats format,
			CancellationToken cancellationToken = default
		)
		{
			var outFiles = new ConcurrentBag<string>();

			void conversion(int index)
			{
				cancellationToken.ThrowIfCancellationRequested();

				var sourceFile = sourceFiles[index];
				var fileName = Path.GetFileNameWithoutExtension(sourceFile);
				var outOneFile = Path.Combine(outFolder, $"{fileName}.{format}");
				
				using var presentation = new Presentation(sourceFile);

				switch (format)
				{
					case SlidesConversionFormats.odp:
					case SlidesConversionFormats.otp:
					case SlidesConversionFormats.pptx:
					case SlidesConversionFormats.pptm:
					case SlidesConversionFormats.potx:
					case SlidesConversionFormats.ppt:
					case SlidesConversionFormats.pps:
					case SlidesConversionFormats.ppsm:
					case SlidesConversionFormats.pot:
					case SlidesConversionFormats.potm:
					case SlidesConversionFormats.pdf:
					case SlidesConversionFormats.xps:
					case SlidesConversionFormats.ppsx:
					case SlidesConversionFormats.tiff:
					case SlidesConversionFormats.html:
					case SlidesConversionFormats.swf:
						{
							var slidesFormat = format.ToString().ParseEnum<SaveFormat>();

							cancellationToken.ThrowIfCancellationRequested();
							presentation.Save(outOneFile, slidesFormat);
							outFiles.Add(outOneFile);
							break;
						}

					case SlidesConversionFormats.txt:
						{
							var lines = new List<string>();
							foreach (var slide in presentation.Slides)
							{
								foreach (var shp in slide.Shapes)
								{
									if (shp is AutoShape ashp)
										lines.Add(ashp.TextFrame.Text);

									cancellationToken.ThrowIfCancellationRequested();
								}

								var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

								if (!string.IsNullOrEmpty(notes))
									lines.Add(notes);

								cancellationToken.ThrowIfCancellationRequested();
							}

							cancellationToken.ThrowIfCancellationRequested();
							System.IO.File.WriteAllLines(outOneFile, lines);
							outFiles.Add(outOneFile);
							break;
						}

					case SlidesConversionFormats.doc:
					case SlidesConversionFormats.docx:
						{
							using (var stream = new MemoryStream())
							{
								cancellationToken.ThrowIfCancellationRequested();
								presentation.Save(stream, SaveFormat.Html);
								stream.Flush();
								stream.Seek(0, SeekOrigin.Begin);

								var doc = new Words.Document(stream);
								cancellationToken.ThrowIfCancellationRequested();
								switch (format)
								{
									case SlidesConversionFormats.doc:
										doc.Save(outOneFile, Words.SaveFormat.Doc);
										break;

									case SlidesConversionFormats.docx:
										doc.Save(outOneFile, Words.SaveFormat.Docx);
										break;

									default:
										throw new ArgumentException($"Unknown format {format}");
								}
							}
							outFiles.Add(outOneFile);
							break;
						}

					case SlidesConversionFormats.bmp:
					case SlidesConversionFormats.jpeg:
					case SlidesConversionFormats.png:
					case SlidesConversionFormats.emf:
					case SlidesConversionFormats.wmf:
					case SlidesConversionFormats.exif:
					case SlidesConversionFormats.ico:
						{
							for (var i = 0; i < presentation.Slides.Count; i++)
							{
								var slide = presentation.Slides[i];
								var outFile = Path.Combine(outFolder, $"{i}.{format}");
								using (var bitmap = slide.GetThumbnail(1, 1))// (new Size((int)size.Width, (int)size.Height)))
								{
									cancellationToken.ThrowIfCancellationRequested();
									bitmap.Save(outFile, format.GetImageFormat());
								}

								outFiles.Add(outFile);
								cancellationToken.ThrowIfCancellationRequested();
							}

							break;
						}

					case SlidesConversionFormats.svg:
						{
							var svgOptions = new SVGOptions
							{
								PicturesCompression = PicturesCompression.DocumentResolution
							};

							for (var i = 0; i < presentation.Slides.Count; i++)
							{
								var slide = presentation.Slides[i];
								var outFile = Path.Combine(outFolder, $"{i}.{format}");
								using (var stream = new FileStream(outFile, FileMode.CreateNew))
								{
									cancellationToken.ThrowIfCancellationRequested();
									slide.WriteAsSvg(stream, svgOptions);
								}

								outFiles.Add(outFile);
								cancellationToken.ThrowIfCancellationRequested();
							}

							break;
						}

					case SlidesConversionFormats.gif:
						{
							_gifEncoder.Encode(presentation, outOneFile, cancellationToken);
							outFiles.Add(outOneFile);
							
							break;
						}

					default:
						throw new ArgumentException($"Unknown format {format}");

				} 
			}

			try
			{
				Parallel.For(0, sourceFiles.Count, conversion);
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}

			return outFiles;
		}

		/// <summary>
		/// Asynchronously converts source file into target format, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file paths.</returns>
		public async Task<IEnumerable<string>> ConversionAsync(
			IList<string> sourceFiles,
			string outFolder,
			SlidesConversionFormats format,
			CancellationToken cancellationToken = default
		) => await Task.Run(() => Conversion(sourceFiles, outFolder, format, cancellationToken));
	}
}
