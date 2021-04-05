using System.IO;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Presentation template provider.
	/// </summary>
	public interface IPresentationTemplateService
	{
		/// <summary>
		/// Returns the template file by the given template identifier (filename).
		/// The result stream is have to dispose outside (on the calling side)!
		/// </summary>
		/// <param name="template">Template identifier.</param>
		/// <param name="cancellationToken">Cancellation token</param>
		/// <returns>The result stream is have to dispose outside (on the calling side)!</returns>
		Task<Stream> GetTemplateStreamAsync(string template, CancellationToken cancellationToken = default);
	}
}
