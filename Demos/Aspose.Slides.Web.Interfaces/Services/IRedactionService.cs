using Aspose.Slides.Web.Interfaces.Models.Redaction;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides redaction logic.
	/// </summary>
	public interface IRedactionService
	{
		/// <summary>
		/// Search for specified string using regular expressions inside source file, replace string with replacement text, saves resulted file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Redaction options.</param>
		/// <returns>Found original lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		string[] Redaction(
			string sourceFile,
			string outFile,
			RedactionOptions options,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Asynchronously search for specified string using regular expressions inside source file, replace string with replacement text, saves resulted file to out file.
		/// Returns found lines.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file.</param>
		/// <param name="options">Redaction options.</param>
		/// <returns>Found original lines. Null if query is invalid.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task<string[]> RedactionAsync(
		   string sourceFile,
		   string outFile,
		   RedactionOptions options,
		   CancellationToken cancellationToken = default
	   );
	}
}
