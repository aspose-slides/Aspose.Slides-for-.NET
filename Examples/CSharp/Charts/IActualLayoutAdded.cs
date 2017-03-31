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
    public class IActualLayoutadded
    {
        public static void Run()
        {
            //ExStart:IActualLayoutadded
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Creating empty presentation
                 using (Presentation pres = new Presentation())
{
               Chart chart = (Chart)pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 500, 350);
               chart.ValidateChartLayout();

               double x = chart.PlotArea.ActualX;
               double y = chart.PlotArea.ActualY;
               double w = chart.PlotArea.ActualWidth;
               double h = chart.PlotArea.ActualHeight;
}
            }
            //ExEnd:IActualLayoutadded
        }
    }
