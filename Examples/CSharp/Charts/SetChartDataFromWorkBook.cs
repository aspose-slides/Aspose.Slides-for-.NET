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
            //ExStart:SetChartDataFromWorkBook
         Presentation pres = new Presentation(dataDir+"Test.pptx");

            IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 500, 400);
            chart.ChartData.ChartDataWorkbook.Clear(0);

            Workbook workbook = null;
            try
            {
                workbook = new Aspose.Cells.Workbook("a1.xlsx");
            }
            catch (Exception ex)
            {
                Console.Write(ex);
            }
            MemoryStream mem = new MemoryStream();
            workbook.Save(mem, Aspose.Cells.SaveFormat.Xlsx);

            chart.ChartData.WriteWorkbookStream(mem);

            chart.ChartData.SetRange("Sheet1!$A$1:$B$9");
            IChartSeries series = chart.ChartData.Series[0];
            series.ParentSeriesGroup.IsColorVaried = true;
            pres.Save(dataDir+"response2.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            //ExEnd:SetChartDataFromWorkBook
        }
    }
}