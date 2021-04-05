using Aspose.Slides.Web.Interfaces.Services;
using Aspose.Slides.Charts;
using System;

namespace Aspose.Slides.Web.Core.Services.Charts
{
	/// <summary>
	/// Implementation logic of creating chart builders
	/// </summary>
	internal sealed class ChartBuilderFactory
	{
		public ChartBuilderFactory()
		{
		}

		/// <summary>
		/// Get instance of IChartBuilder
		/// </summary>
		/// <param name="chartType"></param>
		/// <returns></returns>
		internal IChartBuilder GetChartBuilder(ChartType chartType)
		{
			switch (chartType)
			{
				case ChartType.ClusteredColumn:
				case ChartType.ClusteredBar:
				case ChartType.ClusteredCone:
				case ChartType.ClusteredCylinder:
				case ChartType.ClusteredPyramid:
				case ChartType.ClusteredHorizontalCone:
				case ChartType.ClusteredHorizontalCylinder:
				case ChartType.ClusteredHorizontalPyramid:
				case ChartType.ClusteredBar3D:
				case ChartType.ClusteredColumn3D:
				case ChartType.StackedBar:
				case ChartType.StackedBar3D:
				case ChartType.PercentsStackedBar:
				case ChartType.PercentsStackedBar3D:
				case ChartType.StackedColumn:
				case ChartType.StackedColumn3D:
				case ChartType.PercentsStackedColumn:
				case ChartType.PercentsStackedColumn3D:
				case ChartType.Column3D:
				case ChartType.StackedCylinder:
				case ChartType.PercentsStackedCylinder:
				case ChartType.Cylinder3D:
				case ChartType.StackedCone:
				case ChartType.PercentsStackedCone:
				case ChartType.Cone3D:
				case ChartType.StackedPyramid:
				case ChartType.PercentsStackedPyramid:
				case ChartType.Pyramid3D:
				case ChartType.StackedHorizontalCylinder:
				case ChartType.PercentsStackedHorizontalCylinder:
				case ChartType.StackedHorizontalPyramid:
				case ChartType.PercentsStackedHorizontalPyramid:
				case ChartType.StackedHorizontalCone:
				case ChartType.PercentsStackedHorizontalCone:

					{
						return new BaseChartBuilder(chartType);
					}

				case ChartType.ScatterWithMarkers:
				case ChartType.ScatterWithSmoothLines:
				case ChartType.ScatterWithSmoothLinesAndMarkers:
				case ChartType.ScatterWithStraightLines:
				case ChartType.ScatterWithStraightLinesAndMarkers:
					{
						return new ScatterChartBuilder(chartType);
					}

				case ChartType.Pie:
				case ChartType.Pie3D:
				case ChartType.PieOfPie:
				case ChartType.BarOfPie:
				case ChartType.ExplodedPie:
				case ChartType.ExplodedPie3D:
					{
						return new PieChartBuilder(chartType);
					}

				case ChartType.Doughnut:
				case ChartType.ExplodedDoughnut:

					{
						return new DoughnutChartBuilder(chartType);
					}

				case ChartType.Treemap:
					{
						return new TreemapChartBuilder(chartType);
					}

				case ChartType.OpenHighLowClose:
				case ChartType.VolumeOpenHighLowClose:
				case ChartType.HighLowClose:
				case ChartType.VolumeHighLowClose:
					{
						return new StockChartBuilder(chartType);
					}

				case ChartType.BoxAndWhisker:
					{
						return new BoxAndWhiskerChartBuilder(chartType);
					}

				case ChartType.Funnel:
					{
						return new FunnelChartBuilder(chartType);
					}

				case ChartType.Sunburst:
					{
						return new SunburstChartBuilder(chartType);
					}

				case ChartType.Histogram:
					{
						return new HistogramChartBuilder(chartType);
					}

				case ChartType.ParetoLine:
					{
						return new ParetoLineChartBuilder(chartType);
					}

				case ChartType.Area:
				case ChartType.Area3D:
				case ChartType.PercentsStackedArea:
				case ChartType.PercentsStackedArea3D:
				case ChartType.StackedArea:
				case ChartType.StackedArea3D:
					{
						return new AreaChartBuilder(chartType);
					}

				case ChartType.Bubble:
				case ChartType.BubbleWith3D:
					{
						return new BubbleChartBuilder(chartType);
					}

				case ChartType.Line:
				case ChartType.Line3D:
				case ChartType.LineWithMarkers:
				case ChartType.PercentsStackedLine:
				case ChartType.PercentsStackedLineWithMarkers:
				case ChartType.StackedLine:
				case ChartType.StackedLineWithMarkers:
					{
						return new LineChartBuilder(chartType);
					}

				case ChartType.Radar:
				case ChartType.RadarWithMarkers:
				case ChartType.FilledRadar:
					{
						return new RadarChartBuilder(chartType);
					}

				case ChartType.Surface3D:
				case ChartType.WireframeSurface3D:
				case ChartType.Contour:
				case ChartType.WireframeContour:
					{
						return new SurfaceChartBuilder(chartType);
					}

				case ChartType.Waterfall:
					{
						return new WaterfallChartBuilder(chartType);
					}

				default:
					{
						throw new NotSupportedException($"The char type: {chartType} not support!");
					}
			}
		}
	}
}
