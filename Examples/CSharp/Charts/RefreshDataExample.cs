using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to refresh the chart data by reloading it from the source workbook.
 */

namespace CSharp.Charts
{
    public class RefreshDataExample
    {
        public static void Run()
        {
            // Source file name
            string sourcePath = Path.Combine(RunExamples.GetDataDir_Charts(), "Presentation.pptx");

            using (Presentation pres = new Presentation(sourcePath))
            {
                var chart = (IChart)pres.Slides[0].Shapes[0];
                //refresh the chart data
                chart.ChartData.RefreshData();

                //refresh the data for all charts in a presentation
                pres.RefreshChartsData();
            }
        }
    }
}