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
This example shows how to reduce the resolution of pictures in the presentation and deletes their cropped areas, 
similar to PowerPoint’s “Compress Pictures” applied to all pictures.
 */

namespace Aspose.Slides.Examples.CSharp.Presentations.LowCode
{
    class CompressImages
    {
        public static void Run()
        {
            // Path to source presentation
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_Conversion(), "demo.pptx");

            // The path to output files
            string outFile = Path.Combine(RunExamples.OutPath, "DemoConvertedImage.pptx");

            Console.WriteLine("Presentation size before convertion = {0:F3} MB", new FileInfo(pptxFileName).Length / (1024.0 * 1024.0));
            using (Presentation pres = new Presentation(pptxFileName))
            {
                Aspose.Slides.LowCode.Compress.CompressImages(pres, PicturesCompression.Dpi150);

                pres.Save(outFile, SaveFormat.Pptx);
            }
            Console.WriteLine("Presentation size after convertion = {0:F3} MB", new FileInfo(outFile).Length / (1024.0 * 1024.0));
        }
    }
}