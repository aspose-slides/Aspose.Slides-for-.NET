using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The example shows how to use the StopPreviousSound property of the Effect class to specify whether the animation effect stops the previous sound.
*/

namespace CSharp.Slides.Media
{
    class StopPreviousSoundExample
    {
        public static void Run()
        {
            string pptxFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "AnimationStopSound.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "AnimationStopSound-out.pptx");

            using (Presentation pres = new Presentation(pptxFile))
            {
                // Gets the first effect of the first slide.
                IEffect firstSlideEffect = pres.Slides[0].Timeline.MainSequence[0];

                // Gets the first effect of the second slide.
                IEffect secondSlideEffect = pres.Slides[1].Timeline.MainSequence[0];

                if (firstSlideEffect.Sound != null)
                {
                    // Changes the second effect Enhancements/Sound to "Stop Previous Sound"
                    secondSlideEffect.StopPreviousSound = true;
                }
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
