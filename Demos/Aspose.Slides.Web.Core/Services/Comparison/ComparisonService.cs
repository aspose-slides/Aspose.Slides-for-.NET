using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using Microsoft.Extensions.Logging;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services.Comparison
{
	/// <summary>
	/// Implementation of Comparison business logic.
	/// </summary>
	internal sealed class ComparisonService : SlidesServiceBase, IComparisonService
	{
		private const string partOfDiffFileName = "compared to"; // TODO: need to localize
		
		/// <summary>
		/// Ctor 
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="licenseProvider"></param>
		public ComparisonService(ILogger<ComparisonService> logger, ILicenseProvider licenseProvider) : base(logger)
		{
			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
			licenseProvider.SetAsposeLicense(AsposeProducts.Words);
		}

		/// <summary>
		/// The message for identical presentations
		/// </summary>
		public string MessageForFilesAreIdentical => "The files were compared and they are identical."; // TODO: need to localize

		/// <summary>
		/// Compares two presentations and returns a string name of diff file.
		/// </summary>
		/// <param name="firstPresentationFile">A string path of the first presentation file</param>
		/// <param name="secondPresentationFile">A string path of the second presentation file</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="diffFileSaveFormat">The save format for diff file</param>
		/// <param name="comparisonMethod">The comparison method</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Output file name of diff file</returns>
		public string ComparePresentations(string firstPresentationFile, string secondPresentationFile, string outputPath, ComparisonDiffFileSaveFormats diffFileSaveFormat = ComparisonDiffFileSaveFormats.Pdf, ComparisonMethods comparisonMethod = ComparisonMethods.BySlides, CancellationToken cancellationToken = default)
		{
			if(String.IsNullOrWhiteSpace(firstPresentationFile))
			{
				throw new ArgumentNullException(nameof(firstPresentationFile));
			}

			if(String.IsNullOrWhiteSpace(secondPresentationFile))
			{
				throw new ArgumentNullException(nameof(secondPresentationFile));
			}

			if (String.IsNullOrWhiteSpace(outputPath))
			{
				throw new ArgumentNullException(nameof(outputPath));
			}

			if (firstPresentationFile.Equals(secondPresentationFile))
			{
				throw new ArgumentOutOfRangeException($"{nameof(firstPresentationFile)} name is same as {nameof(secondPresentationFile)} name!");				
			}

			cancellationToken.ThrowIfCancellationRequested();

			_logger.LogDebug($"The first presentation file name: {firstPresentationFile}");
			_logger.LogDebug($"The second presentation file name: {secondPresentationFile}");
			_logger.LogDebug($"The comparison method: {comparisonMethod}");

			var comparator = GetPresentationComparator(comparisonMethod);
			var diffText = comparator.ComparePresentations(firstPresentationFile, secondPresentationFile, cancellationToken);
			var diffContent = String.IsNullOrWhiteSpace(diffText) ? MessageForFilesAreIdentical : diffText;
			
			cancellationToken.ThrowIfCancellationRequested();

			// save diff file
			var diffFileName = $"{outputPath}{Path.DirectorySeparatorChar}{Path.GetFileName(firstPresentationFile)} {partOfDiffFileName} {Path.GetFileName(secondPresentationFile)}";
			var fullDiffFileName = SaveDiffFile(diffContent, diffFileSaveFormat, diffFileName);

			cancellationToken.ThrowIfCancellationRequested();

			_logger.LogInformation($"The: {firstPresentationFile} and {secondPresentationFile} were compared successfully. Result in the: {fullDiffFileName}");

			return fullDiffFileName;
		}

		private IPresentationComparator GetPresentationComparator(ComparisonMethods comparisonMethod)
		{
			return new SlidesTextComparator();
		}

		private string SaveDiffFile(string diffContent, ComparisonDiffFileSaveFormats diffFileSaveFormat, string diffFileName)
		{
			var fullDiffFileName = $"{diffFileName}.{diffFileSaveFormat}";

			var diffDoc = new Words.Document();
			var docBuilder = new Words.DocumentBuilder(diffDoc);

			docBuilder.Writeln(diffContent);

			var outputFormat = diffFileSaveFormat == ComparisonDiffFileSaveFormats.DocX
				? Words.SaveFormat.Docx
				: Words.SaveFormat.Pdf;

			diffDoc.Save(fullDiffFileName, outputFormat);

			return fullDiffFileName;
		}

		/// <summary>
		/// Compares two presentations and returns a string name of diff file asynchronously.
		/// </summary>
		/// <param name="firstPresentationFile">A string path of the first presentation file</param>
		/// <param name="secondPresentationFile">A string path of the second presentation file</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="diffFileSaveFormat">The save format for diff file</param>
		/// <param name="comparisonMethod">The comparison method</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Output file name of diff file</returns>
		public async Task<string> ComparePresentationsAsync(
			string firstPresentationFile,
			string secondPresentationFile,
			string outputPath,
			ComparisonDiffFileSaveFormats diffFileSaveFormat = ComparisonDiffFileSaveFormats.Pdf,
			ComparisonMethods comparisonMethod = ComparisonMethods.BySlides,
			CancellationToken cancellationToken = default)
		=> await Task.Run(() => ComparePresentations(firstPresentationFile, secondPresentationFile, outputPath, diffFileSaveFormat, comparisonMethod, cancellationToken));
	}
}
