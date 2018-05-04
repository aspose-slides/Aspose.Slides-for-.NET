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
	public class FunnelChart
	{

		//ExStart:FunnelChart
		public static void Run()

		{
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Funnel, 50, 50, 500, 400);
				chart.ChartData.Categories.Clear();
				chart.ChartData.Series.Clear();

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				wb.Clear(0);

				chart.ChartData.Categories.Add(wb.GetCell(0, "A1", "Category 1"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A2", "Category 2"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A3", "Category 3"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A4", "Category 4"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A5", "Category 5"));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A6", "Category 6"));

				IChartSeries series = chart.ChartData.Series.Add(ChartType.Funnel);

				series.DataPoints.AddDataPointForFunnelSeries(wb.GetCell(0, "B1", 50));
				series.DataPoints.AddDataPointForFunnelSeries(wb.GetCell(0, "B2", 100));
				series.DataPoints.AddDataPointForFunnelSeries(wb.GetCell(0, "B3", 200));
				series.DataPoints.AddDataPointForFunnelSeries(wb.GetCell(0, "B4", 300));
				series.DataPoints.AddDataPointForFunnelSeries(wb.GetCell(0, "B5", 400));
				series.DataPoints.AddDataPointForFunnelSeries(wb.GetCell(0, "B6", 500));

				pres.Save(dataDir+"Funnel.pptx", SaveFormat.Pptx);
         }

		}
		//ExEnd:FunnelChart
	}
}
