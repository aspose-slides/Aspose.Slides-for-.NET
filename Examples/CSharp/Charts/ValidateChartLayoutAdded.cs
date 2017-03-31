using System.IO;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class ValidateChartLayoutAdded
    {
        public static void Run()
        {
            //ExStart:ValidateChartLayoutAdded
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation pres = new Presentation(dataDir+"test.pptx"))
            {
                Chart chart = (Chart)pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 500, 350);
                chart.ValidateChartLayout();
                double x = chart.PlotArea.ActualX;
                double y = chart.PlotArea.ActualY;
                double w = chart.PlotArea.ActualWidth;
                double h = chart.PlotArea.ActualHeight;
            }
          

            // Saving presentation
            pres.Save(dataDir + "Result.pptx", SaveFormat.Pptx);
            //ExEnd:ValidateChartLayoutAdded
        }
    }
}