using Aspose.Slides.Live.Demos.UI.Helpers;
using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.Util;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using SidesExport = Aspose.Slides.Export;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Converts source file into target format, saves resulted file to out file.
		/// Returns null in case of multiple files (all saved into outFolder).
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>
		/// <returns>Result file path.</returns>
		public string Conversion(
			string sourceFile,
			string outFolder,
			SlidesConversionFormat format
		)
		{
			using (var presentation = new Presentation(sourceFile))
			{
				var fileName = Path.GetFileNameWithoutExtension(sourceFile);
				var outOneFile = Path.Combine(outFolder, $"{fileName}.{format}");

				switch (format)
				{
					case SlidesConversionFormat.odp:
					case SlidesConversionFormat.otp:
					case SlidesConversionFormat.pptx:
					case SlidesConversionFormat.pptm:
					case SlidesConversionFormat.potx:
					case SlidesConversionFormat.ppt:
					case SlidesConversionFormat.pps:
					case SlidesConversionFormat.ppsm:
					case SlidesConversionFormat.pot:
					case SlidesConversionFormat.potm:
					case SlidesConversionFormat.pdf:
					case SlidesConversionFormat.xps:
					case SlidesConversionFormat.ppsx:
					case SlidesConversionFormat.tiff:
					case SlidesConversionFormat.html:
					case SlidesConversionFormat.swf:
						var slidesFormat = format.ToString().ParseEnum<SaveFormat>();
						presentation.Save(outOneFile, slidesFormat);
						return outOneFile;

					case SlidesConversionFormat.txt:
						var lines = new List<string>();
						foreach (var slide in presentation.Slides)
						{
							foreach (var shp in slide.Shapes)
							{
								if (shp is AutoShape ashp)
									lines.Add(ashp.TextFrame.Text);
							}

							var notes = slide.NotesSlideManager.NotesSlide?.NotesTextFrame?.Text;

							if (!string.IsNullOrEmpty(notes))
								lines.Add(notes);
						}
						System.IO.File.WriteAllLines(outOneFile, lines);
						return outOneFile;

					case SlidesConversionFormat.doc:
					case SlidesConversionFormat.docx:
						using (var stream = new MemoryStream())
						{
							presentation.Save(stream, SaveFormat.Html);
							stream.Flush();
							stream.Seek(0, SeekOrigin.Begin);

							var doc = new Words.Document(stream);
							switch (format)
							{
								case SlidesConversionFormat.doc:
									doc.Save(outOneFile, Words.SaveFormat.Doc);
									break;

								case SlidesConversionFormat.docx:
									doc.Save(outOneFile, Words.SaveFormat.Docx);
									break;

								default:
									throw new ArgumentException($"Unknown format {format}");
							}
						}
						return outOneFile;

					case SlidesConversionFormat.bmp:
					case SlidesConversionFormat.jpeg:
					case SlidesConversionFormat.png:
					case SlidesConversionFormat.emf:
					case SlidesConversionFormat.wmf:
					case SlidesConversionFormat.gif:
					case SlidesConversionFormat.exif:
					case SlidesConversionFormat.ico:
						ImageFormat GetImageFormat(SlidesConversionFormat f)
						{
							switch (format)
							{
								case SlidesConversionFormat.bmp:
									return ImageFormat.Bmp;
								case SlidesConversionFormat.jpeg:
									return ImageFormat.Jpeg;
								case SlidesConversionFormat.png:
									return ImageFormat.Png;
								case SlidesConversionFormat.emf:
									return ImageFormat.Wmf;
								case SlidesConversionFormat.wmf:
									return ImageFormat.Wmf;
								case SlidesConversionFormat.gif:
									return ImageFormat.Gif;
								case SlidesConversionFormat.exif:
									return ImageFormat.Emf;
								case SlidesConversionFormat.ico:
									return ImageFormat.Icon;
								default:
									throw new ArgumentException($"Unknown format {format}");
							}
						}

						///var size = presentation.SlideSize.Size;

						for (var i = 0; i < presentation.Slides.Count; i++)
						{
							var slide = presentation.Slides[i];
							var outFile = Path.Combine(outFolder, $"{i}.{format}");
							using (var bitmap = slide.GetThumbnail(1, 1))// (new Size((int)size.Width, (int)size.Height)))
								bitmap.Save(outFile, GetImageFormat(format));
						}

						return null;

					case SlidesConversionFormat.svg:
						var svgOptions = new SVGOptions
						{
							PicturesCompression = PicturesCompression.DocumentResolution
						};

						for (var i = 0; i < presentation.Slides.Count; i++)
						{
							var slide = presentation.Slides[i];
							var outFile = Path.Combine(outFolder, $"{i}.{format}");
							using (var stream = new FileStream(outFile, FileMode.CreateNew))
								slide.WriteAsSvg(stream, svgOptions);
						}

						return null;

					default:
						throw new ArgumentException($"Unknown format {format}");
				}
			}
		}

		/// <summary>
		/// Asynchronously converts source file into target format, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>		
		public string ConvertFile(
			string sourceFile,
			string outFolder,
			SlidesConversionFormat format
		) =>  Conversion(sourceFile, outFolder, format);
	}
}
