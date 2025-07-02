using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.Ink;

/*
The following code sample demonstrates how to use the InkEffect property to get or set Ink effects.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class InkEffectsExample
    {
        public static void Run()
        {
            // The path to the documents directory
            string dataDir = RunExamples.GetDataDir_Shapes();

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "InkEffects.png");

            using (Presentation pres = new Presentation(dataDir + "InkEffects.pptx"))
            {
                // Get Ink object
                Ink.Ink ink = pres.Slides[0].Shapes[0] as Ink.Ink;
                IInkBrush brush = ink.Traces[0].Brush;

                // Show InkEffects of the brush
                Console.WriteLine("InkEffects = {0}", brush.InkEffect);

                // Set image for InkEffects
                IImage image = Images.FromFile(dataDir + "Effect.png");
                Ink.Ink.InkEffectImages.Add(brush.InkEffect, image);

                // Save result
                pres.Slides[0].GetImage(2f, 2f).Save(outFilePath, ImageFormat.Png);
            }
        }
    }
}
