using Aspose.Cells;
using Aspose.Cells.Utility;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Web.API.Clients.Enums;
using Aspose.Slides.Web.Core.Helpers;
using Aspose.Slides.Web.Core.Enums;
using Aspose.Slides.Web.Core.Infrastructure;
using Aspose.Slides.Web.Interfaces.Services;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Microsoft.Extensions.Logging;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	/// <summary>
	/// Implementation business logic of the charts
	/// </summary>
	internal sealed class ChartsService : SlidesServiceBase, IChartsService
	{
		//private const string RowDefaultTitle = "Category";
		//private const string ColumnDefaultTitle = "Series";
		private const float DefaultPreviewWidth = 300f;
		private const float DefaultPreviewHeight = 300f;

		private readonly ChartBuilderFactory _chartBuilderFactory;
		private readonly IConversionService _conversionService;

		/// <summary>
		/// Ctor
		/// </summary>
		/// <param name="logger"></param>
		/// <param name="chartBuilderFactory"></param>
		/// <param name="conversionService"></param>
		/// <param name="licenseProvider"></param>
		public ChartsService(ILogger<ChartsService> logger,
			ChartBuilderFactory chartBuilderFactory,
			IConversionService conversionService,
			ILicenseProvider licenseProvider) : base(logger)
		{
			_chartBuilderFactory = chartBuilderFactory;
			_conversionService = conversionService;

			licenseProvider.SetAsposeLicense(AsposeProducts.Slides);
			licenseProvider.SetAsposeLicense(AsposeProducts.Cells);
		}

		/// <summary>
		/// Creates chart by selected ChartType and from sourceFileData. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceFileData">An excel file with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		/// <returns>Output File Names for Chart</returns>
		public IEnumerable<string> CreateChart(
			ChartTypes chartType,
			string sourceFileData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			bool isPreview = false,
			CancellationToken cancellationToken = default)
		{
			if (String.IsNullOrWhiteSpace(sourceFileData))
			{
				throw new ArgumentNullException(nameof(sourceFileData));
			}

			if (!File.Exists(sourceFileData))
			{
				throw new FileNotFoundException($"The {sourceFileData} not found!");
			}

			if (String.IsNullOrWhiteSpace(outputPath))
			{
				throw new ArgumentNullException(nameof(outputPath));
			}

			var fileExtension = Path.GetExtension(sourceFileData).TrimStart(new char[] { '.' });

			fileExtension = ReplaceOnValidFileFormat(fileExtension);
			var fileType = (FileFormatType)Enum.Parse(typeof(FileFormatType), fileExtension, true);

			if (fileType != FileFormatType.Excel97To2003 &&
				fileType != FileFormatType.Xlsx)
			{
				throw new ArgumentException($"Unknown format {fileType}");
			}

			cancellationToken.ThrowIfCancellationRequested();

			var workbook = new Workbook(sourceFileData);
			using var presentation = new Presentation();

			if(isPreview)
			{
				presentation.SlideSize.SetSize(DefaultPreviewWidth, DefaultPreviewHeight, SlideSizeScaleType.DoNotScale);
			}

			// Aspose.Slides doesn't support .xls for charts
			if (fileType == FileFormatType.Excel97To2003)
			{
				workbook.Save(sourceFileData, SaveFormat.Xlsx); 
			}

			AddChartForWorksheets(presentation, workbook.Worksheets, chartType.ConvertEnum(), sourceFileData, cancellationToken);

			cancellationToken.ThrowIfCancellationRequested();

			// Saving
			var sourceFileName = Path.GetFileNameWithoutExtension(sourceFileData);
			var outputFile = $"{outputPath}{Path.DirectorySeparatorChar}{sourceFileName}.{Export.SaveFormat.Ppt}";

			presentation.Save(outputFile, Export.SaveFormat.Ppt);

			if( !saveFormat.ToString().Equals(Export.SaveFormat.Ppt.ToString(), StringComparison.InvariantCultureIgnoreCase))
			{
				return _conversionService.Conversion(new string[] { outputFile }, outputPath, saveFormat, cancellationToken);
			}

			cancellationToken.ThrowIfCancellationRequested();

			return new List<string>() { outputFile };
		}

		private void AddChartForWorksheets(Presentation presentation, IEnumerable<Worksheet> worksheets, ChartType chartType, string sourceFileData, CancellationToken cancellationToken = default)
		{
			using var sourceFileStream = new FileStream(sourceFileData, FileMode.Open);
			using var sourceMemoryStream = new MemoryStream();

			sourceFileStream.CopyTo(sourceMemoryStream);

			var chartBuilder = _chartBuilderFactory.GetChartBuilder(chartType);
			var slideCount = 0;

			foreach (var worksheet in worksheets)
			{
				if (slideCount != 0)
				{
					presentation.Slides.AddEmptySlide(presentation.Slides[0].LayoutSlide);
				}

				var slide = presentation.Slides[slideCount];

				chartBuilder.CreateChartForWorksheet(worksheet, slide, sourceMemoryStream);
				slideCount++;
				cancellationToken.ThrowIfCancellationRequested();
			}
		}		

		private string ReplaceOnValidFileFormat(string fileExtension)
		{
			switch (fileExtension)
			{
				case "xls":
				{
					return FileFormatType.Excel97To2003.ToString();
				}

				default:
				{
					return fileExtension;
				}
			}
		}

		/// <summary>
		/// Creates chart by selected ChartType and from sourceFileData asynchronously. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceFileData">An excel file with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="cancellationToken"></param>
		/// <returns>Output File Names for Chart</returns>
		public async Task<IEnumerable<string>> CreateChartAsync(
			ChartTypes chartType,
			string sourceFileData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			bool isPreview = false,
			CancellationToken cancellationToken = default)
			=> await Task.Run(() => CreateChart(chartType, sourceFileData, saveFormat, outputPath, isPreview, cancellationToken));

		/// <summary>
		/// Creates a chart by selected ChartType and from sourceJsonData. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceJsonData">A json table (multidimensional array) with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="outputFileName">The output file name</param>
		/// <param name="cancellationToken">A cancellation token to observe while waiting for the task to complete.</param>		
		/// <returns>Output file names for chart</returns>
		public IEnumerable<string> CreateChart(ChartTypes chartType,
			string sourceJsonData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			string outputFileName,
			bool isPreview = false,
			CancellationToken cancellationToken = default)
		{
			if(String.IsNullOrWhiteSpace(sourceJsonData))
			{
				throw new ArgumentNullException(nameof(sourceJsonData));
			}

			if (String.IsNullOrWhiteSpace(outputPath))
			{
				throw new ArgumentNullException(nameof(outputPath));
			}

			if (String.IsNullOrWhiteSpace(outputFileName))
			{
				throw new ArgumentNullException(nameof(outputFileName));
			}							

			cancellationToken.ThrowIfCancellationRequested();

			using var workbook = new Workbook();
						
			CreateWorksheetForJson(workbook, sourceJsonData);

			var sourceFileName = $"{outputPath}{Path.DirectorySeparatorChar}{outputFileName}.{SaveFormat.Xlsx}";
			
			workbook.Save(sourceFileName, SaveFormat.Xlsx);

			using var presentation = new Presentation();

			if (isPreview)
			{
				presentation.SlideSize.SetSize(DefaultPreviewWidth, DefaultPreviewHeight, SlideSizeScaleType.DoNotScale);
			}

			AddChartForWorksheets(presentation, workbook.Worksheets, chartType.ConvertEnum(), sourceFileName, cancellationToken);
			workbook.Dispose();
			cancellationToken.ThrowIfCancellationRequested();

			// save charts
			var outputFile = $"{outputPath}{Path.DirectorySeparatorChar}{outputFileName}.{Aspose.Slides.Export.SaveFormat.Ppt}";

			presentation.Save(outputFile, Aspose.Slides.Export.SaveFormat.Ppt);

			if (!saveFormat.ToString().Equals(Aspose.Slides.Export.SaveFormat.Ppt.ToString(), StringComparison.InvariantCultureIgnoreCase))
			{
				return _conversionService.Conversion(new string[] { outputFile }, outputPath, saveFormat, cancellationToken);
			}

			cancellationToken.ThrowIfCancellationRequested();

			return new List<string>() { outputFile };
		}

		private void CreateWorksheetForJson(Workbook workbook, string jsonInput)
		{
			string firstPropertyName;
			try
			{
				using var jsonDocument = JsonDocument.Parse(jsonInput, new JsonDocumentOptions()
				{
					AllowTrailingCommas = true,
					
				});
				var firstProperty = jsonDocument.RootElement.EnumerateObject().FirstOrDefault();
				firstPropertyName = firstProperty.Name;
			}
			catch (Exception ex) when (ex is InvalidOperationException || ex is JsonException)
			{
				throw new ArgumentNullException(nameof(jsonInput), ex);
			}

			var worksheet = workbook.Worksheets[0];
			worksheet.Name = firstPropertyName;
			
			// Set JsonLayoutOptions
			var options = new JsonLayoutOptions();

			options.ArrayAsTable = true;
			options.IgnoreArrayTitle = true;
			options.ConvertNumericOrDate = true;

			// Import JSON Data
			JsonUtility.ImportData(jsonInput, worksheet.Cells, 0, 0, options);
		}

		/// <summary>
		/// Creates a chart by selected ChartType and from sourceJsonData. 
		/// </summary>
		/// <param name="chartType">The type of chart</param>
		/// <param name="sourceJsonData">A json table (multidimensional array) with data for chart</param>
		/// <param name="saveFormat">Format for saving of chart</param>
		/// <param name="outputPath">The output directory for saving result</param>
		/// <param name="outputFileName">The output file name</param>
		/// <param name="cancellationToken"></param>
		/// <returns>Output file names for chart</returns>
		public async Task<IEnumerable<string>> CreateChartAsync(
			ChartTypes chartType,
			string sourceJsonData,
			SlidesConversionFormats saveFormat,
			string outputPath,
			string outputFileName,
			bool isPreview = false,
			CancellationToken cancellationToken = default)
		=> await Task.Run(() => CreateChart(chartType, sourceJsonData, saveFormat, outputPath, outputFileName, isPreview, cancellationToken));
	}
}
