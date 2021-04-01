using Aspose.Slides.Web.API.Clients.Enums;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for Comparison business logic.
	/// </summary>
	public interface IComparisonService
	{
		/// <summary>
		/// The message for identical presentations
		/// </summary>
		string MessageForFilesAreIdentical { get; }

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
		string ComparePresentations(string firstPresentationFile, string secondPresentationFile, string outputPath, ComparisonDiffFileSaveFormats diffFileSaveFormat = ComparisonDiffFileSaveFormats.Pdf, ComparisonMethods comparisonMethod = ComparisonMethods.BySlides, CancellationToken cancellationToken = default);

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
		Task<string> ComparePresentationsAsync(string firstPresentationFile, string secondPresentationFile, string outputPath, ComparisonDiffFileSaveFormats diffFileSaveFormat = ComparisonDiffFileSaveFormats.Pdf, ComparisonMethods comparisonMethod = ComparisonMethods.BySlides, CancellationToken cancellationToken = default);
	}
}
