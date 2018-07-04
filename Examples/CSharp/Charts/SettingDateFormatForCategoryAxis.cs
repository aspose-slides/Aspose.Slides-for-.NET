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
	class SettingDateFormatForCategoryAxis
	{
		public static void Run()
		{
			//ExStart:SettingDateFormatForCategoryAxis
			// The path to the documents directory.
			string dataDir = RunExamples.GetDataDir_Charts();
			using (Presentation pres = new Presentation())
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Area, 50, 50, 450, 300);

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				wb.Clear(0);

				chart.ChartData.Categories.Clear();
				chart.ChartData.Series.Clear();
				chart.ChartData.Categories.Add(wb.GetCell(0, "A2", new DateTime(2015, 1, 1).ToOADate()));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A3", new DateTime(2016, 1, 1).ToOADate()));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A4", new DateTime(2017, 1, 1).ToOADate()));
				chart.ChartData.Categories.Add(wb.GetCell(0, "A5", new DateTime(2018, 1, 1).ToOADate()));

				IChartSeries series = chart.ChartData.Series.Add(ChartType.Line);
				series.DataPoints.AddDataPointForLineSeries(wb.GetCell(0, "B2", 1));
				series.DataPoints.AddDataPointForLineSeries(wb.GetCell(0, "B3", 2));
				series.DataPoints.AddDataPointForLineSeries(wb.GetCell(0, "B4", 3));
				series.DataPoints.AddDataPointForLineSeries(wb.GetCell(0, "B5", 4));
				chart.Axes.HorizontalAxis.CategoryAxisType = CategoryAxisType.Date;
				chart.Axes.HorizontalAxis.IsNumberFormatLinkedToSource = false;
				chart.Axes.HorizontalAxis.NumberFormat = "yyyy";
				pres.Save(dataDir+"test.pptx", SaveFormat.Pptx);
			}
			//ExEnd:SettingDateFormatForCategoryAxis
		}
	}
}