using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
	public class HistogramChart
	{

		//ExStart:HistogramChart
		public static void Run()

		{

			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Histogram, 50, 50, 500, 400);
				chart.ChartData.Categories.Clear();
				chart.ChartData.Series.Clear();

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				wb.Clear(0);

				IChartSeries series = chart.ChartData.Series.Add(ChartType.Histogram);
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A1", 15));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A2", -41));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A3", 16));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A4", 10));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A5", -23));
				series.DataPoints.AddDataPointForHistogramSeries(wb.GetCell(0, "A6", 16));

				chart.Axes.HorizontalAxis.AggregationType = AxisAggregationType.Automatic;

				pres.Save(dataDir+"Histogram.pptx", SaveFormat.Pptx);
			}

		}

		//ExEnd:HistogramChart
	}
}
