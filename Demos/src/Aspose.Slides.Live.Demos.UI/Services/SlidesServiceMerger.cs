using Aspose.Slides.Live.Demos.UI.Models;
using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace Aspose.Slides.Live.Demos.UI.Services
{
	public partial class SlidesService
	{
		/// <summary>
		/// Merge documents into one file, saves resulted file to out file with specified format.		
		/// </summary>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>
		/// <param name="styleMasterFile">Master file for style in result file. When null, style not changed from source files.</param>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <returns>Result file path.</returns>
		public string Merger(
			string outFolder,
			SlidesConversionFormat format,
			string styleMasterFile,
			params string[] sourceFiles
		)
		{
			var files = new List<string>();
			foreach (var sourceFile in sourceFiles)
			{
				if (Directory.Exists(sourceFile))
					files.AddRange(Directory.EnumerateFiles(sourceFile));
				else
					files.Add(sourceFile);
			}

			var fileName = Path.GetFileNameWithoutExtension(styleMasterFile ?? files.First());
			var resultFile = Path.Combine(outFolder, $"{fileName}.pptx");

			using (var resultPresentation = new Presentation())
			{
				var slides = resultPresentation.Slides;
				while (slides.Any())
					slides.Remove(slides.First());

				var masters = resultPresentation.Masters;

				Presentation masterPresentation = null;
				IMasterSlide masterSlide = null;
				float maxWidth = 0, maxHeight = 0;

				if (styleMasterFile != null)
				{
					masterPresentation = new Presentation(styleMasterFile);
					foreach (var master in masterPresentation.Masters)
						masterSlide = masters.AddClone(master);
				}

				try
				{
					foreach (var sourceFile in files)
					{
						using (var sourcePresentation = new Presentation(sourceFile))
						{
							var sourceSize = sourcePresentation.SlideSize.Size;
							maxWidth = Math.Max(maxWidth, sourceSize.Width);
							maxHeight = Math.Max(maxHeight, sourceSize.Height);

							var master = sourcePresentation.Masters.FirstOrDefault();
							masters.AddClone(master);

							var layout = sourcePresentation.LayoutSlides;
							foreach (var slide in sourcePresentation.Slides)
							{
								if (masterSlide != null)
									slides.AddClone(slide, masterSlide, true);
								else
									slides.AddClone(slide);
							}
						}
					}

					if (styleMasterFile != null)
					{
						var masterSize = masterPresentation.SlideSize.Size;
						resultPresentation.SlideSize.SetSize(masterSize.Width, masterSize.Height, SlideSizeScaleType.DoNotScale);
					}
					else
						resultPresentation.SlideSize.SetSize(maxWidth, maxHeight, SlideSizeScaleType.DoNotScale);
				}
				finally
				{
					masterPresentation?.Dispose();
				}
				resultPresentation.Save(resultFile, Aspose.Slides.Export.SaveFormat.Pptx);
			}

			if (format == SlidesConversionFormat.pptx)
				return resultFile;
			else
				return Conversion(
					resultFile,
					outFolder,
					format
				);
		}

		///// <summary>
		///// Asynchronously merge documents into one file, saves resulted file to out file with specified format.		
		///// </summary>
		///// <param name="outFolder">Output folder.</param>
		///// <param name="format">Output format.</param>
		///// <param name="styleMasterFile">Master file for style in result file. When null, style not changed from source files.</param>
		///// <param name="sourceFiles">Source slides files to proceed.</param>
		///// <returns>Result file path.</returns>
		//public async Task<string> MergerAsync(
		//	string outFolder,
		//	SlidesConversionFormat format,
		//	string styleMasterFile,
		//	params string[] sourceFiles
		//) => await Task.Run(() => Merger(outFolder, format, styleMasterFile, sourceFiles));
	}
}
