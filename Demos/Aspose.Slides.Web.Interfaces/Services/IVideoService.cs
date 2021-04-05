using Aspose.Slides.Web.API.Clients.Enums;
using System.Threading;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for video converter
	/// </summary>
	public interface IVideoService
	{
		/// <summary>
		/// Encodes presentation to the video
		/// </summary>
		/// <param name="sourceFile">The path to the presentation file.</param>
		/// <param name="outFolder">Path to the output folder.</param>
		/// <param name="splitRange">Split range string.</param>
		/// <param name="transitionTime">Slide transition time in seconds.</param>
		/// <param name="codec">Codec type <see cref="VideoCodecs"/></param>
		/// <param name="cancellationToken">Cancellation token</param>
		/// <returns>Returns path to the converted video file.</returns>
		string Encode(
			string sourceFile,
			string outFolder,
			string splitRange,
			int transitionTime,
			VideoCodecs codec,
			CancellationToken cancellationToken);
	}
}
