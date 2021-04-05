using Aspose.Cells;
using Aspose.Slides;
using Aspose.Slides.Charts;
using System.IO;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class BubbleChartBuilder : BaseChartBuilder
	{
		private const float DefaultBubbleChartStartPointX = 0f;
		private const float DefaultBubbleChartStartPointY = 0f;

		public BubbleChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height;

			return slide.Shapes.AddChart(_chartType, DefaultBubbleChartStartPointX, DefaultBubbleChartStartPointY, chartWidth, chartHeight);
		}

		public void CreateChartForWorksheet2(Worksheet worksheet, ISlide slide, MemoryStream memoryStream)
		{
			var chart = GetChart(slide);

			// Setting chart Title			
			chart.ChartTitle.AddTextFrameForOverriding(worksheet.Name);
			chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
			chart.ChartTitle.Height = 20;
			chart.HasTitle = true;

			var rowCount = worksheet.Cells.MaxDataRow + 1;			

			for (var rowIndex = 0; rowIndex < rowCount; rowIndex++)
			{
				var row = worksheet.Cells.Rows[rowIndex];
				var cellX = row.GetCellOrNull(0);
				var cellY = row.GetCellOrNull(1);
				var cellSize = row.GetCellOrNull(2);

				if(cellX?.Type != CellValueType.IsNumeric &&
					cellY?.Type != CellValueType.IsNumeric &&
					cellSize?.Type != CellValueType.IsNumeric)
				{
					continue;
				}

				var series = chart.ChartData.Series.Add(_chartType);

				series.DataPoints.AddDataPointForBubbleSeries(cellX.DoubleValue, cellY.DoubleValue, cellSize.DoubleValue);				
				series.Format.Fill.FillType = FillType.Solid;
				series.Labels.DefaultDataLabelFormat.ShowValue = true;								
			}
		}
	}
}
