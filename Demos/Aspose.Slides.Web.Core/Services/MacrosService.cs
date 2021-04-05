using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;

namespace Aspose.Slides.Web.Core.Services
{
	/// <summary>
	/// The implementation of managing macros logic.
	/// </summary>
	public sealed class MacrosService : SlidesServiceBase, IMacrosService
	{
		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public MacrosService(ILogger<MacrosService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
		}

		/// <summary>
		/// Removes macros from files
		/// </summary>
		/// <param name="sourceFiles">The source files set</param>
		/// <param name="outputDirectory">The output directory for saving result</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>The processed files set</returns>
		public IEnumerable<string> RemoveMacros(IEnumerable<string> sourceFiles, string outputDirectory, CancellationToken cancellationToken = default)
		{
			if(sourceFiles == null || !sourceFiles.Any())
			{
				throw new ArgumentNullException(nameof(sourceFiles));
			}

			if(string.IsNullOrWhiteSpace(outputDirectory))
			{
				throw new ArgumentNullException(nameof(outputDirectory));
			}

			cancellationToken.ThrowIfCancellationRequested();
			IEnumerable<string> resultFiles = null;
			try
			{
				resultFiles = sourceFiles.AsParallel().Select(file => RemoveMacrosFromFile(file, outputDirectory, cancellationToken)).ToArray();
			}
			catch (AggregateException ae)
			{
				foreach (var e in ae.InnerExceptions)
				{
					throw e;
				}
			}

			return resultFiles;
		}

		private string RemoveMacrosFromFile(string sourceFile, string outputDirectory, CancellationToken cancellationToken = default)
		{
			if (!File.Exists(sourceFile))
			{
				throw new FileNotFoundException(nameof(sourceFile));
			}

			cancellationToken.ThrowIfCancellationRequested();

			if (IsPresentation(sourceFile))
			{
				using var presentation = new Presentation(sourceFile);

				presentation.VbaProject?.Modules?.Remove(presentation.VbaProject?.Modules?[0]);

				var resultFileName = Path.GetFileName(sourceFile);
				var resultFullFileName = Path.Combine(outputDirectory, resultFileName);
				var saveFormat = sourceFile.GetSlidesExportSaveFormatBySourceFile();

				cancellationToken.ThrowIfCancellationRequested();

				presentation.Save(resultFullFileName, saveFormat);

				return resultFullFileName;
			}

			return sourceFile;
		}

		private bool IsPresentation(string sourceFile)
		{
			var fileformat = Path.GetExtension(sourceFile).ToSlidesConversionFormats();

			switch(fileformat) 
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
					return true;

				default:
					return false;
			}
		}
	}
}
