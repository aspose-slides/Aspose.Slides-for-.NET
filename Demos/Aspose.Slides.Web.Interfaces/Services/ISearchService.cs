using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for search logic.
	/// </summary>
	public interface ISearchService
	{
		/// <summary>
		/// Search for specified string using regular expressions inside source file, saves found result file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="query">Search query string.</param>
		/// <returns>Found lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		string[] Search(string sourceFile, string query, CancellationToken cancellationToken = default);

		/// <summary>
		/// Asynchronously search for specified string using regular expressions inside source file, saves found result file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="query">Search query string.</param>
		/// <returns>Found lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		Task<string[]> SearchAsync(string sourceFile, string query, CancellationToken cancellationToken = default);
	}
}
