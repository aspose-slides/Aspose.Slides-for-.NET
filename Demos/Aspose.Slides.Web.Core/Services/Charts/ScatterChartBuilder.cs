using Aspose.Slides;
using Aspose.Slides.Charts;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	internal sealed class ScatterChartBuilder : BaseChartBuilder
	{
		private const float DefaultScatterChartWidth = 400f;
		private const float DefaultScatterChartHeight = 400f;

		public ScatterChartBuilder(ChartType chartType) : base(chartType)
		{
		}

		protected override IChart GetChart(ISlide slide)
		{
			var slideSize = slide.Presentation.SlideSize.Size;
			var chartWidth = slideSize.Width <= 0 ? DefaultScatterChartWidth : slideSize.Width;
			var chartHeight = slideSize.Height <= 0 ? DefaultScatterChartHeight : slideSize.Height;

			return slide.Shapes.AddChart(_chartType, DefaultChartStartPointX, DefaultChartStartPointY, chartWidth, chartHeight);
		}		
	}
}
