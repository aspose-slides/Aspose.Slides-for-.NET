using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to determine what parent’s children collection this data point applies to.
*/
namespace CSharp.Charts
{
    class ChartDataPointIndex
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Charts();
            string pptxFile = Path.Combine(dataDir, "ChartIndex.pptx");

            using (Presentation presentation = new Presentation(pptxFile))
            {
                Chart chart = (Chart)presentation.Slides[0].Shapes[0];
                foreach (IChartDataPoint dataPoint in chart.ChartData.Series[0].DataPoints)
                {
                    Console.WriteLine($"Point with index {dataPoint.Index} is applied to {dataPoint.Value}");
                }
            }
        }
    }
}
