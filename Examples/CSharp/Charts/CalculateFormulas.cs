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
This example demonstrates a functionality of an explicit formulas calculation within the workbook.
*/
namespace CSharp.Charts
{
    public class CalculateFormulas
    {
        public static void Run()
        {
            string resultPath = Path.Combine(RunExamples.OutPath, "CalculateFormulas_out.pptx");

            using (Presentation presentation = new Presentation())
            {
                IChart s_chart = presentation.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 10, 10, 600, 300);

                var workbook = s_chart.ChartData.ChartDataWorkbook;
                IChartDataCell cell = workbook.GetCell(0, "A1");
                cell.Formula = "ABS(A2) + MAX(B2:C2)";

                workbook.GetCell(0, "A2").Value = -1;
                workbook.CalculateFormulas();

                workbook.GetCell(0, "B2").Formula = "2";
                workbook.CalculateFormulas();

                workbook.GetCell(0, "C2").Formula = "A2 + 4";
                workbook.CalculateFormulas();

                cell.Formula = "MAX(2:2)";
                workbook.CalculateFormulas();

                presentation.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
