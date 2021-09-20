using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates using the IChartDataWorksheetCollection interface, ChartDataWorksheetCollection class, and IChartDataWorkbook.Worksheets property.
*/

namespace CSharp.Charts
{
    public class WorksheetsExample
    {
        public static void Run()
        {
            //string resultPath = Path.Combine(RunExamples.OutPath, "WorksheetExample.pptx");

            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 400, 500);

                IChartDataWorkbook workbook = chart.ChartData.ChartDataWorkbook;
                for (int i = 0; i < workbook.Worksheets.Count; i++)
                {
                    Console.WriteLine(workbook.Worksheets[i].Name);
                }
            }
        }
    }
}
