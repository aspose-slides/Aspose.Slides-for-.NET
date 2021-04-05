using Aspose.Slides.Web.API.Clients.Enums;
using System.Threading;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides splitter logic.
	/// </summary>
	public interface ISplitterService
	{
		/// <summary>
		/// Splits presentation to parts and saves each part to the specified format.
		/// </summary>
		/// <param name="source">The source presentation file.</param>
		/// <param name="outputDir">Output directory where parts will be stored.</param>
		/// <param name="format">The required format.</param>
		/// <param name="splitType">Splitting type <see cref="SplitTypes"/></param>
		/// <param name="splitNumber">The number of slides in the group (applied only for <see cref="SplitTypes.Number"/>)</param>
		/// <param name="splitRange">The slide ranges string (applied only for <see cref="SplitTypes.Range"/>)</param>
		/// <param name="cancellationToken">The cancellation token.</param>
		void Split(string source,
			string outputDir,
			SlidesConversionFormats format,
			SplitTypes splitType,
			int? splitNumber,
			string splitRange,
			CancellationToken cancellationToken = default);
	}
}
