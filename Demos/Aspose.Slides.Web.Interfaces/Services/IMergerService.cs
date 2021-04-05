using Aspose.Slides.Web.API.Clients.Enums;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for merge business logic.
	/// </summary>
	public interface IMergerService
	{
		/// <summary>
		/// Merge documents into one file, saves resulted file to out file with specified format.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outputFolder">Output folder.</param>
		/// <param name="outputFormat">Output format.</param>
		/// <param name="masterFile">Master file for style in result file. When null, style not changed from source files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result files paths.</returns>
		IEnumerable<string> Merger(
			IEnumerable<string> sourceFiles,
			string outputFolder,
			SlidesConversionFormats outputFormat,
			string masterFile,
			CancellationToken cancellationToken = default);

		/// <summary>
		/// Asynchronously merge documents into one file, saves resulted file to out file with specified format.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outputFolder">Output folder.</param>
		/// <param name="outputFormat">Output format.</param>
		/// <param name="masterFile">Master file for style in result file. When null, style not changed from source files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result files paths.</returns>
		Task<IEnumerable<string>> MergerAsync(
			IEnumerable<string> sourceFiles,
			string outputFolder,
			SlidesConversionFormats outputFormat,
			string masterFile,
			CancellationToken cancellationToken = default);
	}
}
