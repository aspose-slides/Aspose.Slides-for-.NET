using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Charts
{
    class CreateExternalWorkbook
    {
        public static void Run() {

            //ExStart:CreateExternalWorkbook
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();
            using (Presentation pres = new Presentation(dataDir + "presentation.pptx"))
            {
                string externalWbPath = dataDir + "externalWorkbook1.xlsx";

                IChart chart = pres.Slides[0].Shapes.AddChart(ChartType.Pie, 50, 50, 400, 600);

                if (File.Exists(externalWbPath))
                    File.Delete(externalWbPath);

                using (FileStream fileStream = new FileStream(externalWbPath, FileMode.CreateNew))
                {
                    byte[] worbookData = chart.ChartData.ReadWorkbookStream().ToArray();
                    fileStream.Write(worbookData, 0, worbookData.Length);
                }

                chart.ChartData.SetExternalWorkbook(externalWbPath);

                pres.Save(dataDir + "Presentation_with_externalWbPath.pptx", SaveFormat.Pptx);
            }

            //ExEnd:CreateExternalWorkbook

        }
    }
}
