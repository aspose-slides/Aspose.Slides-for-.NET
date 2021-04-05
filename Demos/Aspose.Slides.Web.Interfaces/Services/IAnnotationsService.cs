using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for Annotations service logic.
	/// </summary>
	public interface IAnnotationsService
	{
		/// <summary>
		/// Removes annotations from source file, saves resulted file to out file.
		/// Returns commentaries from file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file. If value is null file not saved.</param>
		/// <returns>Commentaries.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		string[] RemoveAnnotations(string sourceFile, string outFile, CancellationToken cancellationToken = default);

		/// <summary>
		/// Asynchronously removes annotations from source file, saves resulted file to out file.
		/// Returns commentaries from file.
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="outFile">Output slides file. If value is null file not saved.</param>
		/// <returns>Commentaries.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task<string[]> RemoveAnnotationsAsync(string sourceFile, string outFile, CancellationToken cancellationToken = default);		
	}
}
