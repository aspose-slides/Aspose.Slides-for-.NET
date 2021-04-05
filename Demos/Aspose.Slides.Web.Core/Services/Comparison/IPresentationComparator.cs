using System.Threading;

namespace Aspose.Slides.Web.Core.Services.Comparison
{
	/// <summary>
	/// The abstraction for strategy of compare
	/// </summary>
	public interface IPresentationComparator
	{
		/// <summary>
		/// Compares two presentations and returns a string with differents.
		/// </summary>
		/// <param name="firstPresentationFile">A string path of the first presentation file</param>
		/// <param name="secondPresentationFile">A string path of the second presentation file</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>String with differents</returns>
		string ComparePresentations(string firstPresentationFile, string secondPresentationFile, CancellationToken cancellationToken = default);
	}
}
