using System.Threading;
using Viewer = Aspose.Slides.Web.Interfaces.Models.Viewer;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides viewer logic.
	/// </summary>
	public interface IViewerService
	{
		/// <summary>
		/// The Marker id
		/// </summary>
		string MarkerId { get; }

		/// <summary>
		/// Copies source presentation to the output folder, generates SVG representation of slides and returns presentation information.
		/// </summary>
		/// <param name="id">The upload identifier.</param>
		/// <param name="sourceFile">The source presentation file path.</param>
		/// <param name="destinationPath">The resulting presentation file path.</param>
		/// <param name="cancellationToken">The cancellation token.</param>
		/// <returns>The slides information.</returns>
		Viewer.PresentationInfo GetViewerInfo(string id, string sourceFile, string destinationPath, CancellationToken cancellationToken = default);
	}
}
