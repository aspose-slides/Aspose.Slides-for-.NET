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
	class InvertIfNegativeForIndividualSeries
	{
       public static void Run()
		{
			//ExStart:InvertIfNegativeForIndividualSeries
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation())
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 600, 400, true);
				IChartSeriesCollection series = chart.ChartData.Series;
				chart.ChartData.Series.Clear();

				series.Add(chart.ChartData.ChartDataWorkbook.GetCell(0, "B1"), chart.Type);
				series[0].DataPoints.AddDataPointForBarSeries(chart.ChartData.ChartDataWorkbook.GetCell(0, "B2", -5));
				series[0].DataPoints.AddDataPointForBarSeries(chart.ChartData.ChartDataWorkbook.GetCell(0, "B3", 3));
				series[0].DataPoints.AddDataPointForBarSeries(chart.ChartData.ChartDataWorkbook.GetCell(0, "B4", -2));
				series[0].DataPoints.AddDataPointForBarSeries(chart.ChartData.ChartDataWorkbook.GetCell(0, "B5", 1));

				series[0].InvertIfNegative = false;

				series[0].DataPoints[2].InvertIfNegative = true;

				pres.Save(dataDir+ "InvertIfNegativeForIndividualSeries.pptx", SaveFormat.Pptx);
			}

		}

		//ExEnd:InvertIfNegativeForIndividualSeries
	}
}