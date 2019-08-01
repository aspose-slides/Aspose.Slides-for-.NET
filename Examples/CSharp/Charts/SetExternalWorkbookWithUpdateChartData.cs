using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
    class SetExternalWorkbookWithUpdateChartData
    {
        public static void Run() {

            //ExStart:SetExternalWorkbookWithUpdateChartData

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            using (Presentation pres = new Presentation())
            {
                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 400, 600, true);
                IChartData chartData = chart.ChartData;

                (chartData as ChartData).SetExternalWorkbook("http://path/doesnt/exists", false);


                pres.Save(dataDir + "SetExternalWorkbookWithUpdateChartData.pptx", SaveFormat.Pptx);
            }

            //ExEnd:SetExternalWorkbookWithUpdateChartData
        }
    }        
}
