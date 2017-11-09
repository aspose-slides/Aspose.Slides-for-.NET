using System.Drawing;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class MultiCategoryChart
    {
        public static void Run()
        {
            //ExStart:MultiCategoryChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            Presentation pres = new Presentation();
            ISlide slide = pres.Slides[0];

            IChart ch = pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 600, 450);
            ch.ChartData.Series.Clear();
            ch.ChartData.Categories.Clear();


            IChartDataWorkbook fact = ch.ChartData.ChartDataWorkbook;
            fact.Clear(0);
            int defaultWorksheetIndex = 0;

            IChartCategory category = ch.ChartData.Categories.Add(fact.GetCell(0, "c2", "A"));
            category.GroupingLevels.SetGroupingItem(1, "Group1");
            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c3", "B"));

            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c4", "C"));
            category.GroupingLevels.SetGroupingItem(1, "Group2");
            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c5", "D"));

            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c6", "E"));
            category.GroupingLevels.SetGroupingItem(1, "Group3");
            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c7", "F"));

            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c8", "G"));
            category.GroupingLevels.SetGroupingItem(1, "Group4");
            category = ch.ChartData.Categories.Add(fact.GetCell(0, "c9", "H"));

            //            Adding Series
            IChartSeries series = ch.ChartData.Series.Add(fact.GetCell(0, "D1", "Series 1"),
                ChartType.ClusteredColumn);

            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D2", 10));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D3", 20));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D4", 30));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D5", 40));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D6", 50));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D7", 60));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D8", 70));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, "D9", 80));
            // Save presentation with chart
            pres.Save(dataDir+"AsposeChart_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            //ExEnd:MultiCategoryChart
        }
    }
}