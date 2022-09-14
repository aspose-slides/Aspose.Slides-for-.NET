using Aspose.Slides.Animation;
using Aspose.Slides.SlideShow;
using Aspose.Slides.Export;

/*
This example shows how to change an effect Timing/Repeat setting to “Until End of Slide” and "Repeat Until Next Click".
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Transitions
{
    public class AnimationRepeatOnSlide
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Transitions();

            // Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "AnimationOnSlide.pptx"))
            {
                // Gets the effects sequence for the first slide
                ISequence effectsSequence = pres.Slides[0].Timeline.MainSequence;

                // Gets the first effect of the main sequence.
                IEffect effect = effectsSequence[0];

                // Changes the effect Timing/Repeat to "Until End of Slide"
                effect.Timing.RepeatUntilEndSlide = true;

                // Changes the effect Timing/Repeat to "Until End of Slide"
                effect.Timing.RepeatUntilNextClick = true;
                // Save presentation
                pres.Save(RunExamples.OutPath + "AnimationOnSlide-out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}