using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Conversion
{
    class CovertToPDFWithProgressUpdate
    {
        public static void Run() {

            //ExStart:CovertToPDFWithProgressUpdate
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            using (Presentation presentation = new Presentation(dataDir + "ConvertToPDF.pptx"))
            {
                ISaveOptions saveOptions = new PdfOptions();
                saveOptions.ProgressCallback = new ExportProgressHandler();
                presentation.Save(dataDir + "ConvertToPDF.pdf", SaveFormat.Pdf, saveOptions);
            }


            //ExEnd:CovertToPDFWithProgressUpdate
        }
    }

    //ExStart:ExportProgressHandler
    class ExportProgressHandler : IProgressCallback
    {
        public void Reporting(double progressValue)
        {
            // Use progress percentage value here
            int progress = Convert.ToInt32(progressValue);
            Console.WriteLine(progress + "% file converted");
        }
    }

    //ExEnd:ExportProgressHandler
}
