using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides.LowCode;

/*
This example shows how to compress embedded fonts in presentation.
*/

namespace CSharp.Text
{
    class EmbeddedFontCompression
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "presWithEmbeddedFonts.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "presWithEmbeddedFonts-out.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Compress embedded fonts
                Compress.CompressEmbeddedFonts(pres);
                // Save result
                pres.Save(outPath, SaveFormat.Pptx);
            }

            // Get source file info
            FileInfo fi = new FileInfo(presentationName);
            Console.WriteLine("Source file size = {0, 10:N0} bytes", fi.Length);
            // Get result file info
            fi = new FileInfo(outPath);
            Console.WriteLine("Result file size = {0, 10:N0} bytes", fi.Length);
        }
    }
}
