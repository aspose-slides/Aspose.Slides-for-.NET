using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

namespace CSharp.Charts
{
    class EditChartDatainExternalWorkbook
    {
        public static void Run() {

            //ExStart:EditChartDatainExternalWorkbook
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation pres = new Presentation(dataDir + "presentation.pptx"))
            {
                IChart chart = pres.Slides[0].Shapes[0] as IChart;
                ChartData chartData = (ChartData)chart.ChartData;
                               

                chartData.Series[0].DataPoints[0].Value.AsCell.Value = 100;
                pres.Save(dataDir + "presentation_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:EditChartDatainExternalWorkbook
        }
    }
}
