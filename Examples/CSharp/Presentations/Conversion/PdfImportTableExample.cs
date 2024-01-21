using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides.Import;

/*
This example demonstrates how to import PDF document with automatically detected and imported its data as a table in Slide.
*/


namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class PdfImportTableExample
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Conversion();

            string pdfFileName = Path.Combine(dataDir, "SimpleTableExample.pdf");
            string resultPath = Path.Combine(RunExamples.OutPath, "SimpleTableExample.pptx");

            //Create presentation
            using (Presentation pres = new Presentation())
            {
                // Open PDF document
                using (Stream stream = new FileStream(pdfFileName, FileMode.Open, FileAccess.Read, FileShare.Read))
                {
                    // Add Slide to presentation besed on PDF data using automatically detection for importing tables
                    pres.Slides.AddFromPdf(stream, new PdfImportOptions { DetectTables = true });
                }

                // Save result
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}