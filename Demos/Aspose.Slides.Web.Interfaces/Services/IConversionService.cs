using Aspose.Slides.Web.API.Clients.Enums;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides conversion logic.
	/// </summary>
	public interface IConversionService
	{
		/// <summary>
		/// Converts source file into target format, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file paths.</returns>
		IEnumerable<string> Conversion(
			IList<string> sourceFiles,
			string outFolder,
			SlidesConversionFormats format,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Asynchronously converts source file into target format, saves resulted file to out file.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFolder">Output folder.</param>
		/// <param name="format">Output format.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>
		/// <returns>Result file paths.</returns>
		Task<IEnumerable<string>> ConversionAsync(
			IList<string> sourceFiles,
			string outFolder,
			SlidesConversionFormats format,
			CancellationToken cancellationToken = default
		);
	}
}
