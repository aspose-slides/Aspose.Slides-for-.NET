using Aspose.Slides;
using Aspose.Slides.Charts;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class AreaChartBuilder : BaseChartBuilder
	{
		private const float DefaultAreaChartStartPointX = 0f;
		private const float DefaultAreaChartStartPointY = 0f;

		public AreaChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultChartHeight : slideSize.Height;

			return slide.Shapes.AddChart(_chartType, DefaultAreaChartStartPointX, DefaultAreaChartStartPointY, chartWidth, chartHeight);
		}
	}
}
