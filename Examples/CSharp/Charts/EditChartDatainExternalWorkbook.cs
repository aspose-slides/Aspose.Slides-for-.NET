using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

namespace CSharp.Charts
{
    class EditChartDatainExternalWorkbook
    {
        public static void Run() {

            // Pay attention the path to external workbook is hardly saved in the presentation
            // so please copy file externalWorkbook.xlsx from Data/Chart directory D:\Aspose.Slides\Aspose.Slides-for-.NET-master\Examples\Data\Charts\ before run the example

            //ExStart:EditChartDatainExternalWorkbook
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation pres = new Presentation(dataDir + "presentation.pptx"))
            {
                IChart chart = pres.Slides[0].Shapes[0] as IChart;
                ChartData chartData = (ChartData)chart.ChartData;
                               

                chartData.Series[0].DataPoints[0].Value.AsCell.Value = 100;
                pres.Save(RunExamples.OutPath + "presentation_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:EditChartDatainExternalWorkbook
        }
    }
}
