using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example how to use PdfOptions.RasterizeUnsupportedFontStyles property, which indicates whether text should be rasterized as a bitmap and saved to PDF 
when the font does not support bold styling. This approach can enhance the quality of text in the resulting PDF for certain fonts.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertToPdfUnsupportedFontStyles
    {
        public static void Run()
        {
            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "UnsupportedFontStyles.pdf");

            // Create a new presentation
            using (Presentation pres = new Presentation())
            {
                // Save the presentation to PDF
                pres.Save(outFilePath, SaveFormat.Pdf, new PdfOptions
                {
                    RasterizeUnsupportedFontStyles = true
                });
            }
        }
    }
}
