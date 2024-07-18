using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Aspose.Slides.Export;

/*
The following example shows how to export a presentation to PDF, including embedded files.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ExportOlekExample
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            // The path to output file.
            string outFilePath = Path.Combine(RunExamples.OutPath, "PresOleExample.pdf");

            using (Presentation pres = new Presentation(dataDir + "PresOleExample.pptx"))
            {
                PdfOptions options = new PdfOptions();
                // Include OLE data into exported PDF.
                options.IncludeOleData = true;
                // Save result
                pres.Save(outFilePath, SaveFormat.Pdf, options);
            }
        }
    }
}
