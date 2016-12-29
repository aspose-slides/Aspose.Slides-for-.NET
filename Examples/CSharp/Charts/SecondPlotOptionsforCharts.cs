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
    public class SecondPlotOptionsforCharts
    {
        public static void Run()
        {
            //ExStart:SecondPlotOptionsforCharts
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            // Add chart on slide
            IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.PieOfPie, 50, 50, 500, 400);
     
            // Set different properties
            chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;
            chart.ChartData.Series[0].ParentSeriesGroup.SecondPieSize = 149;
            chart.ChartData.Series[0].ParentSeriesGroup.PieSplitBy = Aspose.Slides.Charts.PieSplitType.ByPercentage;
            chart.ChartData.Series[0].ParentSeriesGroup.PieSplitPosition = 53;

            // Write presentation to disk
            presentation.Save(dataDir + "SecondPlotOptionsforCharts_out.pptx", SaveFormat.Pptx);
            //ExEnd:SecondPlotOptionsforCharts

        }
    }
}