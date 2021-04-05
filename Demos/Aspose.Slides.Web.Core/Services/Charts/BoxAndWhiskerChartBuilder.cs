using Aspose.Slides;
using Aspose.Slides.Charts;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class BoxAndWhiskerChartBuilder : BaseChartBuilder
	{
		private const float DefaultBoxAndWhiskerChartStartPointX = 0f;
		private const float DefaultBoxAndWhiskerChartStartPointY = 0f;

		public BoxAndWhiskerChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height;

			return slide.Shapes.AddChart(_chartType, DefaultBoxAndWhiskerChartStartPointX, DefaultBoxAndWhiskerChartStartPointY, chartWidth, chartHeight, false);
		}
	}
}
