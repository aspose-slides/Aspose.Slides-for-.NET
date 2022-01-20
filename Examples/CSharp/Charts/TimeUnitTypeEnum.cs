using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This sample demonstrates using  TimeUnitType enumeration
*/

namespace CSharp.Charts
{
    public class TimeUnitTypeEnum
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "TimeUnitTypeEnum.pptx");

            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Area, 10, 10, 400, 300, true);
                chart.Axes.HorizontalAxis.MajorUnitScale = TimeUnitType.None;
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
