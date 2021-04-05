using Aspose.Slides.Web.API.Clients.Enums;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Threading;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Presentation signature service interface.
	/// </summary>
	public interface ISignatureService
	{
		/// <summary>
		/// Adds signature to the each presentation slide.
		/// </summary>
		/// <param name="inputFile">The presentation file.</param>
		/// <param name="destinationFolder">The output folder.</param>
		/// <param name="format">The output format.</param>
		/// <param name="image">The signature image stream. Should be null when the text signature is added.</param>
		/// <param name="text">The signature text.</param>
		/// <param name="color">The color of the signature text (ignored when image signature is added).</param>
		/// <param name="cancellationToken">Cancellation token.</param>
		/// <returns></returns>
		IEnumerable<string> Sign(string inputFile, string destinationFolder, SlidesConversionFormats format, Stream image, string text, Color color, CancellationToken cancellationToken = default);
	}
}
