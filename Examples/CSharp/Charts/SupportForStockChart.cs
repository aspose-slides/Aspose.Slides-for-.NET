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
	class SupportForStockChart
	{
		public static void Run()
		{
			//ExStart:SupportForStockChart
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation(dataDir+"Test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.OpenHighLowClose, 50, 50, 600, 400, false);

				chart.ChartData.Series.Clear();
				chart.ChartData.Categories.Clear();

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				chart.ChartData.Categories.Add(wb.GetCell(0, 1, 0, "A"));
				chart.ChartData.Categories.Add(wb.GetCell(0, 2, 0, "B"));
				chart.ChartData.Categories.Add(wb.GetCell(0, 3, 0, "C"));

				chart.ChartData.Series.Add(wb.GetCell(0, 0, 1, "Open"), chart.Type);
				chart.ChartData.Series.Add(wb.GetCell(0, 0, 2, "High"), chart.Type);
				chart.ChartData.Series.Add(wb.GetCell(0, 0, 3, "Low"), chart.Type);
				chart.ChartData.Series.Add(wb.GetCell(0, 0, 4, "Close"), chart.Type);

				IChartSeries series = chart.ChartData.Series[0];

				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 1, 1, 72));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 2, 1, 25));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 3, 1, 38));

				series = chart.ChartData.Series[1];
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 1, 2, 172));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 2, 2, 57));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 3, 2, 57));

				series = chart.ChartData.Series[2];
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 1, 3, 12));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 2, 3, 12));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 3, 3, 13));

				series = chart.ChartData.Series[3];
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 1, 4, 25));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 2, 4, 38));
				series.DataPoints.AddDataPointForStockSeries(wb.GetCell(0, 3, 4, 50));

				chart.ChartData.SeriesGroups[0].UpDownBars.HasUpDownBars = true;
				chart.ChartData.SeriesGroups[0].HiLowLinesFormat.Line.FillFormat.FillType = FillType.Solid;

				foreach (IChartSeries ser in chart.ChartData.Series)
				{
					ser.Format.Line.FillFormat.FillType = FillType.NoFill;
				}

				pres.Save(dataDir+"output.pptx", SaveFormat.Pptx);
			}

		}
		//ExEnd:SupportForStockChart
	}
}
