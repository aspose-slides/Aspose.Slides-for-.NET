using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Core.Services.Conversion
{
	/// <summary>
	/// Interface of gif encoder logic.
	/// </summary>
	public interface IGifEncoder
	{
		/// <summary>
		/// Encodes to gif format.
		/// </summary>
		/// <param name="presentation"></param>
		/// <param name="outputFileName"></param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns></returns>
		string Encode(IPresentation presentation, string outputFileName, CancellationToken cancellationToken = default);

		/// <summary>
		/// Encodes to gif format asynchronously.
		/// </summary>
		/// <param name="presentation"></param>
		/// <param name="outputFileName"></param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns></returns>
		Task<string> EncodeAsync(IPresentation presentation, string outputFileName, CancellationToken cancellationToken = default);
	}
}
