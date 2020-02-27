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
	public class TreeMapChart
	{

		//ExStart:TreeMapChart
		public static void Run()
		{

         
		string dataDir = RunExamples.GetDataDir_Charts();
           using (Presentation pres = new Presentation(dataDir+"test.pptx"))
			{
				IChart chart = pres.Slides[0].Shapes.AddChart(Aspose.Slides.Charts.ChartType.Treemap, 50, 50, 500, 400);
				chart.ChartData.Categories.Clear();
				chart.ChartData.Series.Clear();

				IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

				wb.Clear(0);

				//branch 1
				IChartCategory leaf = chart.ChartData.Categories.Add(wb.GetCell(0, "C1", "Leaf1"));
				leaf.GroupingLevels.SetGroupingItem(1, "Stem1");
				leaf.GroupingLevels.SetGroupingItem(2, "Branch1");

				chart.ChartData.Categories.Add(wb.GetCell(0, "C2", "Leaf2"));

				leaf = chart.ChartData.Categories.Add(wb.GetCell(0, "C3", "Leaf3"));
				leaf.GroupingLevels.SetGroupingItem(1, "Stem2");

				chart.ChartData.Categories.Add(wb.GetCell(0, "C4", "Leaf4"));


				//branch 2
				leaf = chart.ChartData.Categories.Add(wb.GetCell(0, "C5", "Leaf5"));
				leaf.GroupingLevels.SetGroupingItem(1, "Stem3");
				leaf.GroupingLevels.SetGroupingItem(2, "Branch2");

				chart.ChartData.Categories.Add(wb.GetCell(0, "C6", "Leaf6"));

				leaf = chart.ChartData.Categories.Add(wb.GetCell(0, "C7", "Leaf7"));
				leaf.GroupingLevels.SetGroupingItem(1, "Stem4");

				chart.ChartData.Categories.Add(wb.GetCell(0, "C8", "Leaf8"));

				IChartSeries series = chart.ChartData.Series.Add(Aspose.Slides.Charts.ChartType.Treemap);
				series.Labels.DefaultDataLabelFormat.ShowCategoryName = true;
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D1", 4));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D2", 5));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D3", 3));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D4", 6));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D5", 9));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D6", 9));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D7", 4));
				series.DataPoints.AddDataPointForTreemapSeries(wb.GetCell(0, "D8", 3));

				series.ParentLabelLayout = ParentLabelLayoutType.Overlapping;

				pres.Save("Treemap.pptx", SaveFormat.Pptx);
			}

		}
        //ExEnd:TreeMapChart
	}
}