using Aspose.Slides.Web.Interfaces.Models.Metadata;
using System.Threading;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for metadata logic.
	/// </summary>
	public interface IMetadataService
	{
		/// <summary>
		/// Gets presentation metadata.
		/// </summary>
		/// <param name="sourceFile">Path to the presentation file.</param>
		/// <returns>Metadata object.</returns>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		PresentationMetadata GetMetadata(string sourceFile, CancellationToken cancellationToken = default);

		/// <summary>
		/// Updates presentation metadata.
		/// </summary>
		/// <param name="sourceFile">Path to the source presentation file.</param>
		/// <param name="outFile">Path to the resulting presentation file with applied metadata.</param>
		/// <param name="metadata">Metadata object.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		void UpdateMetadata(
			string sourceFile,
			string outFile,
			PresentationMetadata metadata,
			CancellationToken cancellationToken = default
		);
	}
}
