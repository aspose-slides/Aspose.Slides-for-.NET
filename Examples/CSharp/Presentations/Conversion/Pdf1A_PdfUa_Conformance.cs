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
This example demonstrates saving presentation to PDF/A-1a and PDF/UA compliant document.
*/
namespace CSharp.Presentations.Conversion
{
    class Pdf1A_PdfUa_Conformance
    {
        public static void Run()
        {
            string pptxFile = Path.Combine(RunExamples.GetDataDir_Conversion(), "tagged-pdf-demo.pptx");
            string outPdf1aFile = Path.Combine(RunExamples.OutPath, "tagged-pdf-demo_1a.pdf");
            string outPdf1bFile = Path.Combine(RunExamples.OutPath, "tagged-pdf-demo_1b.pdf");
            string outPdfUaFile = Path.Combine(RunExamples.OutPath, "tagged-pdf-demo_1ua.pdf");

            using (Presentation presentation = new Presentation(pptxFile))
            {
                presentation.Save(outPdf1aFile, SaveFormat.Pdf,
                    new PdfOptions { Compliance = PdfCompliance.PdfA1a });

                presentation.Save(outPdf1bFile, SaveFormat.Pdf,
                    new PdfOptions { Compliance = PdfCompliance.PdfA1b });

                presentation.Save(outPdfUaFile, SaveFormat.Pdf,
                    new PdfOptions { Compliance = PdfCompliance.PdfUa });
            }
        }
    }
}
