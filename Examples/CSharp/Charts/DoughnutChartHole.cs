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
    public class DoughnutChartHole
    {
        public static void Run()
        {
            //ExStart:DoughnutChartHole
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.Doughnut, 50, 50, 400, 400);
            chart.ChartData.SeriesGroups[0].DoughnutHoleSize = 90;

            // Write presentation to disk
            presentation.Save(dataDir + "DoughnutHoleSize_out.pptx", SaveFormat.Pptx);
            //ExEnd:DoughnutChartHole
        }
    }
}