using System.IO;
using Aspose.Slides.Export;

/*
This example imports a PDF document into Presentation. 
A new SlideCollection.AddFromPdf method creates slides from the PDF document 
and adds them to the end of the collection
*/


namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class PdfImportExample
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Conversion();

            string pdfFileName = Path.Combine(dataDir, "welcome-to-powerpoint.pdf");
            string resultPath = Path.Combine(RunExamples.OutPath, "fromPdfDocument.pptx");

            using (Presentation pres = new Presentation())
            {
                pres.Slides.AddFromPdf(pdfFileName);
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}