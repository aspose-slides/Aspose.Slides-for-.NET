using Aspose.Slides.Web.Interfaces.Models.Watermark;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for slides watermark logic.
	/// </summary>
	public interface IWatermarkService
	{
		/// <summary>
		/// Adds text watermark into source files, saves resulted files to out files.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		void AddTextWatermark(
			IList<string> sourceFiles,
			IList<string> outFiles,
			TextWatermarkOptions options,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Asynchronously adds text watermark into source files, saves resulted files to out files.
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task AddTextWatermarkAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			TextWatermarkOptions options,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Adds image watermark into source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFile">Source slides file to proceed.</param>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		void AddImageWatermark(
			IList<string> sourceFiles,
			IList<string> outFiles,
			ImageWatermarkOptions options,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Asynchronously adds image watermark into source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="options">Watermark options.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task AddImageWatermarkAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			ImageWatermarkOptions options,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Removes watermark from source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		void RemoveWatermark(
			IList<string> sourceFiles,
			IList<string> outFiles,
			CancellationToken cancellationToken = default
		);

		/// <summary>
		/// Asynchronously removes watermark from source file, saves resulted file to out file.		
		/// </summary>
		/// <param name="sourceFiles">Source slides files to proceed.</param>
		/// <param name="outFiles">Output slides files.</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		Task RemoveWatermarkAsync(
			IList<string> sourceFiles,
			IList<string> outFiles,
			CancellationToken cancellationToken = default
		);
	}
}
