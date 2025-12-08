using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;
using Aspose.Slides.LowCode;

/*
This example shows how to convert presentations into sets of raster images.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.LowCode
{
    class ConvertToImage
    {
        public static void Run()
        {
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_LowCode(), "ConvertExample.pptx");
            string outPathJpeg = Path.Combine(RunExamples.OutPath, "ConvertedToJpeg.jpg");
            string outPathPng = Path.Combine(RunExamples.OutPath, "ConvertedToPng.png");
            string outPathTiff = Path.Combine(RunExamples.OutPath, "ConvertedToTiff.tiff");

            using (Presentation pres = new Presentation(pptxFileName))
            {
                Aspose.Slides.LowCode.Convert.ToJpeg(pres, outPathJpeg);
                Aspose.Slides.LowCode.Convert.ToPng(pres, outPathPng);
                Aspose.Slides.LowCode.Convert.ToTiff(pres, outPathTiff);

            }
        }
    }
}