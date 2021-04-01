using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.Charts;
using System;
using System.IO;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class StockChartBuilder : BaseChartBuilder
	{
		public StockChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height;

			return slide.Shapes.AddChart(_chartType, DefaultChartStartPointX, DefaultChartStartPointY, chartWidth, chartHeight, false);
		}

		public override void CreateChartForWorksheet(Worksheet worksheet, ISlide slide, MemoryStream memoryStream)
		{
			var chart = GetChart(slide);

			// Setting chart Title			
			chart.ChartTitle.AddTextFrameForOverriding(worksheet.Name);
			chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
			chart.ChartTitle.Height = 20;
			chart.HasTitle = true;

			int defaultWorksheetIndex = 0;

			// Getting the chart data worksheet
			IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

			// Delete default generated series and categories
			chart.ChartData.Series.Clear();
			chart.ChartData.Categories.Clear();

			var rowCount = worksheet.Cells.MaxDataRow + 1;
			var colCount = worksheet.Cells.MaxDataColumn + 1;			

			for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
			{				
				var row = worksheet.Cells.Rows[rowIndex];

				for (var colIndex = 0; colIndex < colCount; colIndex++)
				{					
					var cell = row.GetCellOrNull(colIndex);

					if (rowIndex == 0 && colIndex == 0)
					{						
						continue;
					}

					// Adding new series
					if (rowIndex == 0)
					{						
						if (cell.Type != CellValueType.IsString)
						{
							throw new ArgumentException($"The cell[{rowIndex}:{colIndex}] has incorrect data.");
						}

						var colTitle = cell.GetStringValue(CellValueFormatStrategy.DisplayStyle);						

						chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, colIndex, colTitle), _chartType);						
					}

					// Adding new categories
					if (colIndex == 0)
					{
						if (cell.Type != CellValueType.IsString && cell.Type != CellValueType.IsDateTime)
						{
							throw new ArgumentException($"The cell[{rowIndex}:{colIndex}] has incorrect data.");
						}

						var rowTitle = cell.GetStringValue(CellValueFormatStrategy.DisplayStyle);						

						chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, colIndex, rowTitle));
					}

					if (cell.Type == CellValueType.IsNumeric)
					{						
						var series = chart.ChartData.Series[colIndex - 1];
						series.DataPoints.AddDataPointForStockSeries(fact.GetCell(defaultWorksheetIndex, rowIndex, colIndex, cell.Value));
					}
				}
			}

			chart.ChartData.SeriesGroups[0].UpDownBars.HasUpDownBars = true;
			chart.ChartData.SeriesGroups[0].HiLowLinesFormat.Line.FillFormat.FillType = FillType.Solid;			

			foreach (IChartSeries ser in chart.ChartData.Series)
			{
				ser.Format.Line.FillFormat.FillType = FillType.NoFill;
			}
		}
	}
}
