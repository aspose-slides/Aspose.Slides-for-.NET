using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// Implementation of merge business logic.
	/// </summary>
	internal sealed class MergerService : SlidesServiceBase, IMergerService
	{
		private readonly IConversionService _conversionService;

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="conversionService"></param>
		/// <param name="licenseProvider"></param>
		public MergerService(ILogger<MergerService> logger, IConversionService conversionService, ILicenseProvider licenseProvider) : base(logger)
		{
			_conversionService = conversionService;
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Merge documents into one file, saves resulted file to out file with specified format.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outputFolder">Output folder.</param>
		/// <param name="outputFormat">Output format.</param>
		/// <param name="masterFile">Master file for style in result file. When null, style not changed from source files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result files paths.</returns>
		public IEnumerable<string> Merger(
			IEnumerable<string> sourceFiles,
			string outputFolder,
			SlidesConversionFormats outputFormat,
			string masterFile,
			CancellationToken cancellationToken = default)
		{
			var inputFiles = new List<string>();

			foreach (var sourceFile in sourceFiles)
			{
				if (Directory.Exists(sourceFile))
					inputFiles.AddRange(Directory.EnumerateFiles(sourceFile));
				else
					inputFiles.Add(sourceFile);

				cancellationToken.ThrowIfCancellationRequested();
			}

			var fileName = Path.GetFileNameWithoutExtension(masterFile ?? inputFiles.First());
			var resultFile = Path.Combine(outputFolder, $"{fileName}.pptx");

			cancellationToken.ThrowIfCancellationRequested();

			using var resultPresentation = new Presentation();
			
			cancellationToken.ThrowIfCancellationRequested();

			var resultSlides = resultPresentation.Slides;

			while (resultSlides.Any())
			{
				resultSlides.Remove(resultSlides.First());
				cancellationToken.ThrowIfCancellationRequested();
			}

			var resultMastersSlides = resultPresentation.Masters;

			Presentation masterPresentation = null;
			IMasterSlide masterSlide = null;
			float maxWidth = 0, maxHeight = 0;

			if (masterFile != null)
			{
				masterPresentation = new Presentation(masterFile);

				foreach (var master in masterPresentation.Masters)
				{
					masterSlide = resultMastersSlides.AddClone(master);
					cancellationToken.ThrowIfCancellationRequested();
				}
			}

			try
			{
				foreach (var inputFile in inputFiles)
				{
					using var sourcePresentation = new Presentation(inputFile);
					
					cancellationToken.ThrowIfCancellationRequested();

					var sourceSize = sourcePresentation.SlideSize.Size;

					maxWidth = Math.Max(maxWidth, sourceSize.Width);
					maxHeight = Math.Max(maxHeight, sourceSize.Height);

					var sourceMasterSlide = sourcePresentation.Masters.FirstOrDefault();

					resultMastersSlides.AddClone(sourceMasterSlide);

					foreach (var slide in sourcePresentation.Slides)
					{
						if (masterSlide != null)
							resultSlides.AddClone(slide, masterSlide, true);
						else
							resultSlides.AddClone(slide);

						cancellationToken.ThrowIfCancellationRequested();
					}

					cancellationToken.ThrowIfCancellationRequested();
				}

				if (masterFile != null)
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

			cancellationToken.ThrowIfCancellationRequested();
			resultPresentation.Save(resultFile, Export.SaveFormat.Pptx);

			if (outputFormat == SlidesConversionFormats.pptx)
			{
				return new List<string>() { resultFile };
			}
			
			return _conversionService.Conversion(
					new string[] { resultFile },
					outputFolder,
					outputFormat,
					cancellationToken
				);
		}

		/// <summary>
		/// Asynchronously merge documents into one file, saves resulted file to out file with specified format.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outputFolder">Output folder.</param>
		/// <param name="outputFormat">Output format.</param>
		/// <param name="masterFile">Master file for style in result file. When null, style not changed from source files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result files paths.</returns>
		public Task<IEnumerable<string>> MergerAsync(
			IEnumerable<string> sourceFiles,
			string outputFolder,
			SlidesConversionFormats outputFormat,
			string masterFile,
			CancellationToken cancellationToken = default)
		{
			IEnumerable<string> result = null;
			Exception exception = null;
			var thr = new Thread(
				() =>
				{
					try
					{
						result = Merger(sourceFiles, outputFolder, outputFormat, masterFile, cancellationToken);
					}
					catch (Exception ex)
					{
						_logger.LogError(ex, "Merging error.");
						exception = ex;
					}
				}
			)
			{
				Priority = ThreadPriority.BelowNormal
			};
			thr.Start();
			if (!thr.Join(TimeSpan.FromMinutes(3)))
			{
				thr.Interrupt();

				throw new ProcessingTimeoutException("Processing timeout");
			}
			else
			{
				if (exception != null)
					throw exception;

				return Task.FromResult(result);
			}
		}
	}
}
