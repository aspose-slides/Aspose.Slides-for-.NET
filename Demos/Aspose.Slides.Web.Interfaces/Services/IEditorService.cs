using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides edit logic.
	/// </summary>
	public interface IEditorService
	{
		/// <summary>
		/// Replaces slides with given svg-files in the presentation.
		/// </summary>
		/// <param name="sourceFile">The source presentation file.</param>
		/// <param name="slides">The list of given svg-files.</param>
		/// <param name="outFile">The path to the resulting file</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns></returns>
		Task ReplaceSlidesAsync(
			string sourceFile,
			IEnumerable<string> slides,
			string outFile,
			CancellationToken cancellationToken = default
		);
	}
}
