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
This example shows how to control the compression level of the presentation document.
*/

namespace CSharp.Presentations.Saving
{
    class CompressionLevelExample
    {
        public static void Run()
        {
            // The path to output files
            string outFileLevel1 = Path.Combine(RunExamples.OutPath, "PresentationCompressionLevel1.pptx");
            string outFileLevel9 = Path.Combine(RunExamples.OutPath, "PresentationCompressionLevel9.pptx");

            using (Presentation pres = new Presentation())
            {
                // Fastest compression with the lowest compression ratio.
                pres.Save(outFileLevel1, SaveFormat.Pptx, new PptxOptions
                {
                    CompressionLevel = CompressionLevel.Level1
                });

                // Maximum compression. Produces the smallest file size with the slowest processing speed.
                pres.Save(outFileLevel9, SaveFormat.Pptx, new PptxOptions
                {
                    CompressionLevel = CompressionLevel.Level9
                });
            }
        }
    }
}