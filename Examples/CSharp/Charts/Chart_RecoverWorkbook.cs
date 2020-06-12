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
This example demonstrates how to recover data from chart cache if the data source of the chart is an external workbook and it's not available.
*/
namespace CSharp.Charts
{
    class Chart_RecoverWorkbook
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Charts();

            string pptxFile = Path.Combine(dataDir, "ExternalWB.pptx");
            string outPptxFile = Path.Combine(RunExamples.OutPath, "ExternalWB_out.pptx");

            LoadOptions lo = new LoadOptions();
            lo.SpreadsheetOptions.RecoverWorkbookFromChartCache = true;

            using (Presentation pres = new Presentation(pptxFile, lo))
            {
                IChart chart = pres.Slides[0].Shapes[0] as IChart;
                IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

                pres.Save(outPptxFile, SaveFormat.Pptx);
            }
        }
    }
}
