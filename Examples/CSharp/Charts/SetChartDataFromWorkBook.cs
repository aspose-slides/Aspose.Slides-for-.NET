using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Cells;
using Aspose.Slides.Examples.CSharp;


namespace CSharp.Charts
{
    class SetChartDataFromWorkBook
    {

        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation(/*dataDir + "Test.pptx"*/))
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 500, 400);
                chart.ChartData.ChartDataWorkbook.Clear(0);

                Workbook workbook = null;
                try
                {
                    workbook = new Aspose.Cells.Workbook(dataDir + "book1.xlsx");
                }
                catch (Exception ex)
                {
                    Console.Write(ex);
                }

                MemoryStream mem = new MemoryStream();
                workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

                mem.Position = 0;
                chart.ChartData.WriteWorkbookStream(mem);

                chart.ChartData.SetRange("Sheet2!$A$1:$B$3");
                IChartSeries series = chart.ChartData.Series[0];
                series.ParentSeriesGroup.IsColorVaried = true;
                pres.Save(Path.Combine(RunExamples.OutPath, "response2.pptx"), Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}