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
This example shows how to reduces the size of a presentation by applying all available compression steps: 
removing unused master slides and layout slides, compressing embedded fonts by retaining only the glyphs in use, 
reducing picture resolution to 150 DPI, and removing cropped areas of pictures. 
 */

namespace Aspose.Slides.Examples.CSharp.Presentations.LowCode
{
    class CompressPresentation
    {
        public static void Run()
        {
            // Path to source presentation
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_Conversion(), "demo.pptx");

            // The path to output files
            string outFile = Path.Combine(RunExamples.OutPath, "DemoConverted.pptx");

            Console.WriteLine("Presentation size before convertion = {0:F3} MB", new FileInfo(pptxFileName).Length / (1024.0 * 1024.0));
            using (Presentation pres = new Presentation(pptxFileName))
            {
                Aspose.Slides.LowCode.Compress.CompressPresentation(pres);

                pres.Save(outFile, SaveFormat.Pptx);
            }
            Console.WriteLine("Presentation size after convertion = {0:F3} MB", new FileInfo(outFile).Length / (1024.0 * 1024.0));
        }
    }
}