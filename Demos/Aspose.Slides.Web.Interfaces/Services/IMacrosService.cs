using System.Collections.Generic;
using System.Threading;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for manage macros logic.
	/// </summary>
	public interface IMacrosService
	{
		/// <summary>
		/// Removes macros from files
		/// </summary>
		/// <param name="sourceFiles">The source files set</param>
		/// <param name="outputDirectory">The output directory for saving result</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>The processed files set</returns>
		IEnumerable<string> RemoveMacros(
			IEnumerable<string> sourceFiles,
			string outputDirectory,
			CancellationToken cancellationToken = default);
	}
}
