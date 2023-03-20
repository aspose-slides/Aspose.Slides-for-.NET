using System;
using Aspose.Slides.Animation;
using Aspose.Slides.SlideShow;
using Aspose.Slides.Export;

/*
This example shows how to to specify whether an effect will rewind after playing.
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Transitions
{
    public class AnimationRewind
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Transitions();

            // Instantiate Presentation class that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "AnimationRewind.pptx"))
            {
                // Gets the effects sequence for the first slide
                ISequence effectsSequence = presentation.Slides[0].Timeline.MainSequence;

                // Gets the first effect of the main sequence.
                IEffect effect = effectsSequence[0];
                Console.WriteLine("\nEffect Timing/Rewind in source presentation is {0}", effect.Timing.Rewind);
                // Turns the effect Timing/Rewind on.
                effect.Timing.Rewind = true;

                // Save presentation
                presentation.Save(RunExamples.OutPath + "AnimationRewind-out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);


                using (Presentation pres = new Presentation(RunExamples.OutPath + "AnimationRewind-out.pptx"))
                {
                    // Gets the effects sequence for the first slide
                    effectsSequence = pres.Slides[0].Timeline.MainSequence;

                    // Gets the first effect of the main sequence.
                    effect = effectsSequence[0];
                    Console.WriteLine("Effect Timing/Rewind in destination presentation is {0}\n", effect.Timing.Rewind);
                }
            }
        }
    }
}