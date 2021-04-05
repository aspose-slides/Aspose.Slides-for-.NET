using Aspose.Slides.Web.API.Clients.Enums;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;

namespace Aspose.Slides.Web.Interfaces.Services
{
	/// <summary>
	/// Interface for charts business logic.
	/// </summary>
	public interface IChartsService
	{
		/// <summary>
		/// Creates a chart by selected ChartType and from sourceFileData. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceFileData">An excel file with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="isPreview">Make chart full size or preview</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Output file names for chart</returns>
		IEnumerable<string> CreateChart(
			ChartTypes chartType,
			string sourceFileData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			bool isPreview = false,
			CancellationToken cancellationToken = default);

		/// <summary>
		/// Creates a chart by selected ChartType and from sourceFileData asynchronously. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceFileData">An excel file with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="isPreview">Make chart full size or preview</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Output file names for chart</returns>
		Task<IEnumerable<string>> CreateChartAsync(
			ChartTypes chartType,
			string sourceFileData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			bool isPreview = false,
			CancellationToken cancellationToken = default);

		/// <summary>
		/// Creates a chart by selected ChartType and from sourceJsonData. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceJsonData">A json table (multidimensional array) with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="outputFileName">The output file name</param>
		/// <param name="isPreview">Make chart full size or preview</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Output file names for chart</returns>
		IEnumerable<string> CreateChart(
			ChartTypes chartType,
			string sourceJsonData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			string outputFileName,
			bool isPreview = false,
			CancellationToken cancellationToken = default);

		/// <summary>
		/// Creates a chart by selected ChartType and from sourceJsonData. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceJsonData">A json table (multidimensional array) with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="outputFileName">The output file name</param>
		/// <param name="isPreview">Make chart full size or preview</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete</param>
		/// <returns>Output file names for chart</returns>
		Task<IEnumerable<string>> CreateChartAsync(
			ChartTypes chartType,
			string sourceJsonData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			string outputFileName,
			bool isPreview = false,
			CancellationToken cancellationToken = default);
	}
}
