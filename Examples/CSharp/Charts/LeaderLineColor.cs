using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System.Drawing;
using System.IO;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;

/*
This example shows how to change color of all leader lines in the collection.
*/

namespace CSharp.Charts
{
    class LeaderLineColor
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Charts(), "LeaderLinesColor.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "LeaderLinesColor-out.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Get the chart from the first slide
                IChart chart = (IChart)pres.Slides[0].Shapes[0];

                // Get series of the chart
                IChartSeriesCollection series = chart.ChartData.Series;

                // Get lebels of the first serie
                IDataLabelCollection labels = series[0].Labels;

                // Change color of all leader lines in the collection
                labels.LeaderLinesFormat.Line.FillFormat.SolidFillColor.Color = Color.FromArgb(255, 255, 0, 0);

                // Save result
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
