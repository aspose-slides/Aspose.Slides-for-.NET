using Aspose.Slides;
using Aspose.Slides.Charts;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class HistogramChartBuilder : BaseChartBuilder
	{
		private const float DefaultHistogramChartStartPointX = 0f;
		private const float DefaultHistogramChartStartPointY = 0f;

		public HistogramChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height;

			return slide.Shapes.AddChart(_chartType, DefaultHistogramChartStartPointX, DefaultHistogramChartStartPointY, chartWidth, chartHeight);
		}		
	}
}
