using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example shows how to provides options that control the look of Ink objects in exported documents.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ExportInkExample
    {
        public static void Run()
        {
            // The path to the documents directory
            string dataDir = RunExamples.GetDataDir_Conversion();

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "HideInkDemo.pdf");

            using (Presentation pres = new Presentation(dataDir + "InkOptions.pptx"))
            {
                PdfOptions options = new PdfOptions();
                // Hide ink objects
                options.InkOptions.HideInk = true;
                // Save result
                pres.Save(outFilePath, SaveFormat.Pdf, options);

                // Show Ink objects
                options.InkOptions.HideInk = false;
                // Set using ROP operation for rendering brush
                options.InkOptions.InterpretMaskOpAsOpacity = false;
                // Set path to output file
                outFilePath = Path.Combine(RunExamples.OutPath, "ROPInkDemo.pdf");
                // Save result
                pres.Save(outFilePath, SaveFormat.Pdf, options);
            }
        }
    }
}
