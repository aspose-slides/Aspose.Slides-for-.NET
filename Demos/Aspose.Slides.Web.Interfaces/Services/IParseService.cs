using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides parse logic.
	/// </summary>
	public interface IParseService
	{
		/// <summary>
		/// Parse documents into text and image files, saves resulted file to out folder.
		/// Returns parsed parts details.
		/// </summary>
		/// <param name="outFolder">Output folder file. If value is null files not saved.</param>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <returns>Extracted data details.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		List<(string file, string text, string[] media)> Parser(
			string outFolder,
			CancellationToken cancellationToken = default,
			params string[] sourceFiles
		);

		/// <summary>
		/// Asynchronously parse documents into text and image files, saves resulted file to out folder.
		/// Returns parsed parts details.
		/// </summary>
		/// <param name="outFolder">Output folder file. If value is null files not saved.</param>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <returns>Extracted data details.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task<List<(string file, string text, string[] images)>> ParserAsync(
			string outFolder,
			CancellationToken cancellationToken = default,
			params string[] sourceFiles
		);
	}
}
