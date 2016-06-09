using System;
using System.Collections.Generic;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.ProgrammersGuide.Presentations
{
    class CustomOptionsPDFConversion
    {
        public static void Run()
        {           
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "ConvertToPDF.pptx"))
            {
                //Instantiate the PdfOptions class
                PdfOptions pdfOptions = new PdfOptions();

                // Set Jpeg Quality
                pdfOptions.JpegQuality = 90;

                // Define behavior for metafiles
                pdfOptions.SaveMetafilesAsPng = true;

                // Set Text Compression level
                pdfOptions.TextCompression = PdfTextCompression.Flate;

                // Define the PDF standard
                pdfOptions.Compliance = PdfCompliance.Pdf15;

                // Save the presentation to PDF with specified options
                pres.Save(dataDir + "Custom_Option_Pdf_Conversion.pdf", SaveFormat.Pdf, pdfOptions);
            }
        }
    }
}