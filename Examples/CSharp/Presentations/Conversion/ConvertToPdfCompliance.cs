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
Aspose.Slides supports PDF export to formats 1.6, 1.7
and operations with A2a, A2b, A2u, A3a and A3b compliance levels:
PdfCompliance.PdfA2a
PdfCompliance.PdfA2b
PdfCompliance.PdfA2u
PdfCompliance.PdfA3a
PdfCompliance.PdfA3b

This C# code demonstrates an operation where a PDF is saved based on the PdfA2a compliance level
*/

namespace CSharp.Presentations.Conversion
{
    class ConvertToPdfCompliance
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "ConvertToPDF.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "ConvertToPDF-Comp.pdf");

            using (Presentation presentation = new Presentation(presentationName))
            {
                PdfOptions pdfOptions = new PdfOptions() {Compliance = PdfCompliance.PdfA2a};
                presentation.Save(outPath, SaveFormat.Pdf, pdfOptions);
            }
        }
    }
}
