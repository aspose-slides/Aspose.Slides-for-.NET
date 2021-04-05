using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// The interface of validation logic
	/// </summary>
	public interface IFileValidatorService
	{
		/// <summary>
		/// Validates files
		/// </summary>
		/// <param name="fileName">The input file for validation</param>		
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Is valid or not</returns>
		bool IsValidFile(string fileName, CancellationToken cancellationToken = default);

		/// <summary>
		/// Validates files asynchronously
		/// </summary>
		/// <param name="fileName">The input file for validation</param>		
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Is valid or not</returns>
		Task<bool> IsValidFileAsync(string fileName, CancellationToken cancellationToken = default);
	}
}
