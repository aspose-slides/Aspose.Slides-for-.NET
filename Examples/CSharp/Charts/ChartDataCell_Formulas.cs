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
This example demonstrates a way to set a formula value for a chart data cell.
*/
namespace CSharp.Charts
{
    class ChartDataCell_Formulas
    {
        public static void Run()
        {
            string outpptxFile = Path.Combine(RunExamples.OutPath, "ChartDataCell_Formulas_out.pptx");

            using (Presentation presentation = new Presentation())
            {
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 150, 150, 500, 300);
                IChartDataWorkbook workbook = chart.ChartData.ChartDataWorkbook;

                IChartDataCell cell1 = workbook.GetCell(0, "B2");
                cell1.Formula = "1 + SUM(F2:H5)";

                IChartDataCell cell2 = workbook.GetCell(0, "C2");
                cell2.R1C1Formula = "MAX(R2C6:R5C8) / 3";

                presentation.Save(outpptxFile, SaveFormat.Pptx);
            }
        }
    }
}
