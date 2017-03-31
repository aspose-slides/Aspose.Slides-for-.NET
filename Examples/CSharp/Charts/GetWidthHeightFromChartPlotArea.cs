using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class GetWidthHeightFromChartPlotArea
    {
        public static void Run()
        {
            //ExStart:GetWidthHeightFromChartPlotArea
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation(dataDir+"test.Pptx"))
            {
                Chart chart = (Chart)pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 500, 350);
                chart.ValidateChartLayout();

                double x = chart.PlotArea.ActualX;
                double y = chart.PlotArea.ActualY;
                double w = chart.PlotArea.ActualWidth;
                double h = chart.PlotArea.ActualHeight;
            }
                        
            // Save presentation with chart
            pres.Save(dataDir + "Chart_out.pptx", SaveFormat.Pptx);
            //ExEnd:GetWidthHeightFromChartPlotArea
        }
    }
}