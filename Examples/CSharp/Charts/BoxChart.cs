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
	public class BoxChart
	{
		//ExStart:BoxChart
		public static void Run()
		{
			string dataDir = RunExamples.GetDataDir_Charts();

			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.BoxAndWhisker, 50, 50, 500, 400);
				chart.ChartData.Categories.Clear();
				chart.ChartData.Series.Clear();

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				wb.Clear(0);

				chart.ChartData.Categories.Add(wb.GetCell(0, "A1", "Category 1"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A2", "Category 1"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A3", "Category 1"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A4", "Category 1"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A5", "Category 1"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A6", "Category 1"));

				IChartSeries series = chart.ChartData.Series.Add(ChartType.BoxAndWhisker);

				series.QuartileMethod = QuartileMethodType.Exclusive;
				series.ShowMeanLine = true;
				series.ShowMeanMarkers = true;
				series.ShowInnerPoints = true;
				series.ShowOutlierPoints = true;

				series.DataPoints.AddDataPointForBoxAndWhiskerSeries(wb.GetCell(0, "B1", 15));
				series.DataPoints.AddDataPointForBoxAndWhiskerSeries(wb.GetCell(0, "B2", 41));
				series.DataPoints.AddDataPointForBoxAndWhiskerSeries(wb.GetCell(0, "B3", 16));
				series.DataPoints.AddDataPointForBoxAndWhiskerSeries(wb.GetCell(0, "B4", 10));
				series.DataPoints.AddDataPointForBoxAndWhiskerSeries(wb.GetCell(0, "B5", 23));
				series.DataPoints.AddDataPointForBoxAndWhiskerSeries(wb.GetCell(0, "B6", 16));


				pres.Save("BoxAndWhisker.pptx", SaveFormat.Pptx);
			}


		}
		//ExEnd:BoxChart
	}
}