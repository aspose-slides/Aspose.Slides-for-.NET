using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;

/*
This example demonstrates how to set the duration of the slide transition effect in milliseconds. 
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Transitions
{
    class AnimationDurationSlide
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Transitions();
            string outPath = Path.Combine(RunExamples.OutPath, "AnimationDurationSlidest-out.pptx");

            // Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "AnimationDurationSlides.pptx"))
            {
                foreach (var slide in pres.Slides)
                {
                    // Sets the transition duration to 0.25s
                    slide.SlideShowTransition.Duration = 250;
                }

                // Save result
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
