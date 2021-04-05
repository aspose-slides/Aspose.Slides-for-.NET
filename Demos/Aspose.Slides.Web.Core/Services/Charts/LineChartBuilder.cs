using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.Charts;
using System;
using System.IO;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class LineChartBuilder : BaseChartBuilder
	{
		public LineChartBuilder(ChartType chartType) : base(chartType)
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
			var hasColumnTitle = false;
			var hasRowTitle = false;
			var columnShift = 0;
			var rowShift = 0;
			var indexShift = 1;
			string zeroCellValue = null;

			for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
			{
				var row = worksheet.Cells.Rows[rowIndex];

				for (var colIndex = 0; colIndex < colCount; colIndex++)
				{
					var cell = row.GetCellOrNull(colIndex);

					if (rowIndex == 0 && colIndex == 0 && cell?.Type != CellValueType.IsNumeric)
					{
						hasRowTitle = true;
						indexShift = 0;

						if (cell != null)
						{
							zeroCellValue = cell.GetStringValue(CellValueFormatStrategy.DisplayStyle);
						}

						continue;
					}

					// Adding new categories
					if (rowIndex == 0)
					{
						var colTitle = $"{ColumnDefaultTitle} {colIndex + indexShift}";

						if (cell.Type == CellValueType.IsString)
						{
							colTitle = cell.GetStringValue(CellValueFormatStrategy.DisplayStyle);
							hasColumnTitle = true;
							hasRowTitle = false;
						}

						if (hasColumnTitle && !String.IsNullOrWhiteSpace(zeroCellValue))
						{
							chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, colIndex, zeroCellValue));
							zeroCellValue = null;
							columnShift = 1;
						}

						chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, colIndex + columnShift + indexShift, colTitle));						
					}

					// Adding new series
					if (colIndex == 0)
					{
						var rowTitle = $"{RowDefaultTitle} {rowIndex + indexShift}";

						if (cell.Type == CellValueType.IsString)
						{
							rowTitle = cell.GetStringValue(CellValueFormatStrategy.DisplayStyle);
							hasRowTitle = true;
						}

						if (hasRowTitle && !String.IsNullOrWhiteSpace(zeroCellValue))
						{
							chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, rowIndex, colIndex, zeroCellValue), _chartType);
							zeroCellValue = null;
							rowShift = 1;
						}

						chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, rowIndex + rowShift + indexShift, colIndex, rowTitle), _chartType);
					}

					if (cell.Type == CellValueType.IsNumeric)
					{
						int dataRow = rowIndex + indexShift + rowShift;
						int dataColumn = colIndex + indexShift + columnShift;

						// Now populating series data
						IChartSeries series;

						if (hasRowTitle && hasColumnTitle)
						{
							series = chart.ChartData.Series[rowIndex - 1];
						}
						else if (hasColumnTitle)
						{
							series = chart.ChartData.Series[rowIndex];
						}
						else if (hasRowTitle)
						{
							series = chart.ChartData.Series[rowIndex - 1];
							dataColumn++;
							dataRow++;
						}
						else
						{
							series = chart.ChartData.Series[rowIndex];
						}

						series.DataPoints.AddDataPointForLineSeries(fact.GetCell(defaultWorksheetIndex, dataRow, dataColumn, cell.Value));						
						series.Format.Fill.FillType = FillType.Solid;						
					}
				}
			}
		}
	}
}
