using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
    class SetExternalWorkbook
    {
        public static void Run() {
            //ExStart:SetExternalWorkbook
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 400, 600, false);
                IChartData chartData = chart.ChartData;
                                
                chartData.SetExternalWorkbook(dataDir+ "externalWorkbook.xlsx");
                              

                chartData.Series.Add(chartData.ChartDataWorkbook.GetCell(0, "B1"), ChartType.Pie);
                chartData.Series[0].DataPoints.AddDataPointForPieSeries(chartData.ChartDataWorkbook.GetCell(0, "B2"));
                chartData.Series[0].DataPoints.AddDataPointForPieSeries(chartData.ChartDataWorkbook.GetCell(0, "B3"));
                chartData.Series[0].DataPoints.AddDataPointForPieSeries(chartData.ChartDataWorkbook.GetCell(0, "B4"));

                chartData.Categories.Add(chartData.ChartDataWorkbook.GetCell(0, "A2"));
                chartData.Categories.Add(chartData.ChartDataWorkbook.GetCell(0, "A3"));
                chartData.Categories.Add(chartData.ChartDataWorkbook.GetCell(0, "A4"));
                pres.Save(dataDir + "Presentation_with_externalWorkbook.pptx", SaveFormat.Pptx);
            }

            //ExEnd:SetExternalWorkbook
        }
    }
}
