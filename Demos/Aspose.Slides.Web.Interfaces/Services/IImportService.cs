using Aspose.Slides.Web.API.Clients.Enums;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// The interface for service logic of importing.
	/// </summary>
	public interface IImportService
	{
		/// <summary>
		/// Converts source files into target format of presentation, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source files to proceed</param>
		/// <param name="outputFolder">Output folder</param>
		/// <param name="conversionFormat">Presentation available formats.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file path</returns>
		string ImportToPresentation(
			IEnumerable<string> sourceFiles,
			string outputFolder,
			PresentationFormats conversionFormat,
			CancellationToken cancellationToken = default);


		/// <summary>
		/// Converts source files into target format of presentation, saves resulted file to out file asynchronously.
		/// </summary>
		/// <param name="sourceFiles">Source files to proceed</param>
		/// <param name="outputFolder">Output folder</param>
		/// <param name="conversionFormat">Presentation available formats.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file path</returns>
		Task<string> ImportToPresentationAsync(
			IEnumerable<string> sourceFiles,
			string outputFolder,
			PresentationFormats conversionFormat,
			CancellationToken cancellationToken = default);
	}
}
