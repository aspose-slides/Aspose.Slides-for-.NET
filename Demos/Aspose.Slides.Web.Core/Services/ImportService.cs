using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Drawing;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Aspose.Slides.Web.Core.Infrastructure;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// The implementation of service logic of importing.
	/// </summary>
	internal sealed class ImportService : SlidesServiceBase, IImportService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public ImportService(ILogger<ImportService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Converts source files into target format of presentation, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source files to proceed</param>
		/// <param name="outputFolder">Output folder</param>
		/// <param name="conversionFormat">Presentation available formats.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file path</returns>
		public string ImportToPresentation(IEnumerable<string> sourceFiles, string outputFolder, PresentationFormats conversionFormat, CancellationToken cancellationToken = default)
		{
			if(sourceFiles == null || !sourceFiles.Any())
			{
				throw new ArgumentNullException(nameof(sourceFiles));
			}

			if(String.IsNullOrWhiteSpace(outputFolder))
			{
				throw new ArgumentNullException(nameof(outputFolder));
			}

			using var presentation = new Presentation();
			
			for(int i = 0; i < presentation.Slides.Count; i++)
			{
				presentation.Slides.RemoveAt(i);
			}

			foreach (var sourceFile in sourceFiles)
			{
				cancellationToken.ThrowIfCancellationRequested();
				
				var sourceFileExt = Path.GetExtension(sourceFile);
				var sourceFileFormat = sourceFileExt.TrimStart('.').ToSlidesConversionFormats();

				if(sourceFileFormat == SlidesConversionFormats.pdf)
				{
					presentation.Slides.AddFromPdf(sourceFile);
				}
				else if(IsImage(sourceFileFormat))
				{
					ImportImage(presentation, sourceFile);
				}
				else
				{
					throw new ArgumentException($"The {sourceFile} is not supported file format.");
				}
			}

			cancellationToken.ThrowIfCancellationRequested();

			var outputFile = Path.Combine(outputFolder, $"{Path.GetFileNameWithoutExtension(sourceFiles.First())}.{conversionFormat.ToSaveFormat()}");
			
			presentation.Save(outputFile, conversionFormat.ToSaveFormat());

			return outputFile;
		}

		private static bool IsImage(SlidesConversionFormats fileFormat)
		{
			return fileFormat == SlidesConversionFormats.bmp
			       || fileFormat == SlidesConversionFormats.gif
			       || fileFormat == SlidesConversionFormats.ico
			       || fileFormat == SlidesConversionFormats.jpeg
			       || fileFormat == SlidesConversionFormats.png
			       || fileFormat == SlidesConversionFormats.tiff;
		}

		private void ImportImage(IPresentation presentation, string sourceImage)
		{
			// See https://www.hanselman.com/blog/how-do-you-use-systemdrawing-in-net-core
			using var image = System.Drawing.Image.FromFile(sourceImage, true);
			var presImage = presentation.Images.AddImage(image);
			var slide = CreateSlide(presentation);

			var imageSize = new SizeF(presImage.Width, presImage.Height)
				.ResizeKeepAspect(presentation.SlideSize.Size.Width, presentation.SlideSize.Size.Height);
			slide.Shapes.AddPictureFrame(ShapeType.Rectangle,
				(presentation.SlideSize.Size.Width - imageSize.Width) / 2,
				(presentation.SlideSize.Size.Height - imageSize.Height) / 2,
				imageSize.Width,
				imageSize.Height,
				presImage);
		}

		private static ISlide CreateSlide(IPresentation presentation)
		{
			ISlide slide = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);
			slide.Shapes.Clear();
			return slide;
		}

		/// <summary>
		/// Converts source files into target format of presentation, saves resulted file to out file asynchronously.
		/// </summary>
		/// <param name="sourceFiles">Source files to proceed</param>
		/// <param name="outputFolder">Output folder</param>
		/// <param name="conversionFormat">Presentation available formats.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file path</returns>
		public async Task<string> ImportToPresentationAsync(IEnumerable<string> sourceFiles, string outputFolder, PresentationFormats conversionFormat, CancellationToken cancellationToken = default)
			=> await Task.Run(() => ImportToPresentation(sourceFiles, outputFolder, conversionFormat, cancellationToken));		
	}
}
