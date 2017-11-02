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
    public class SettingAutomicPieChartSliceColors
    {
        public static void Run()
        {
            //ExStart:SettingAutomicPieChartSliceColors
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            // Instantiate Presentation class that represents PPTX file
            using (Presentation presentation = new Presentation())
            {
              // Instantiate Presentation class that represents PPTX file
             Presentation presentation = new Presentation();

             // Access first slide
             ISlide slides = presentation.Slides[0];

             // Add chart with default data
             IChart chart = slides.Shapes.AddChart(ChartType.Pie, 100, 100, 400, 400);

             // Setting chart Title
             chart.ChartTitle.AddTextFrameForOverriding("Sample Title");
             chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
             chart.ChartTitle.Height = 20;
             chart.HasTitle = true;

             // Set first series to Show Values
             chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;

             // Setting the index of chart data sheet
             int defaultWorksheetIndex = 0;

             // Getting the chart data worksheet
             IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

             // Delete default generated series and categories
             chart.ChartData.Series.Clear();
             chart.ChartData.Categories.Clear();

             // Adding new categories
             chart.ChartData.Categories.Add(fact.GetCell(0, 1, 0, "First Qtr"));
             chart.ChartData.Categories.Add(fact.GetCell(0, 2, 0, "2nd Qtr"));
             chart.ChartData.Categories.Add(fact.GetCell(0, 3, 0, "3rd Qtr"));

             // Adding new series
             IChartSeries series = chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, "Series 1"), chart.Type);

             // Now populating series data
             series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 20));
             series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 50));
             series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 30));
            
             series.ParentSeriesGroup.IsColorVaried = true;
             presentation.Save("C:\\Aspose Data\\Pie.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
         }
            }
            //ExEnd:SettingAutomicPieChartSliceColors
        }
    }
