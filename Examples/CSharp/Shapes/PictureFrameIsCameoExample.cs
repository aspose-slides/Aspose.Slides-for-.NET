using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
The following code sample demonstrates how to use the IsCameo property for PictureFrame.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class PictureFrameIsCameoExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Shapes(), "PresCameo.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Check if first picture frame is Cameo
                PictureFrame shape = pres.Slides[0].Shapes[0] as PictureFrame;
                if (shape != null)
                {
                    Console.WriteLine("First picture is Cameo: " + shape.IsCameo);
                }

                // Check if third picture frame is Cameo
                shape = pres.Slides[0].Shapes[2] as PictureFrame;
                if (shape != null)
                {
                    Console.WriteLine("Third picture is Cameo: " + shape.IsCameo);
                }
            }
        }
    }
}
