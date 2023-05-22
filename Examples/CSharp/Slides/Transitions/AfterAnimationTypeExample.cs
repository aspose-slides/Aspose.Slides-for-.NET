using System.Drawing;
using System.IO;
using Aspose.Slides.Animation;
using Aspose.Slides.SlideShow;
using Aspose.Slides.Export;

/*
This example demonstrates how to use Effect.AfterAnimationColor alongside AfterAnimationType.
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Transitions
{
    public class AfterAnimationTypeExample
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Transitions();
            string outPath = Path.Combine(RunExamples.OutPath, "AnimationAfterEffect-out.pptx");

            // Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "AnimationAfterEffect.pptx"))
            {
                // Add new slide to the presentation
                ISlide slide1 = pres.Slides.AddClone(pres.Slides[0]);
                // Get the first effect of the first slide
                ISequence seq = slide1.Timeline.MainSequence;
                // Change the After animation effect to "Hide on Next Mouse Click" 
                foreach (IEffect effect in seq)
                    effect.AfterAnimationType = AfterAnimationType.HideOnNextMouseClick;

                // Add new slide to the presentation
                ISlide slide2 = pres.Slides.AddClone(pres.Slides[0]);
                // Get the first effect of the first slide
                seq = slide2.Timeline.MainSequence;
                // Change the After animation effect type to "Color"
                foreach (IEffect effect in seq)
                {
                    effect.AfterAnimationType = AfterAnimationType.Color;
                    effect.AfterAnimationColor.Color = Color.Green;
                }

                // Add new slide to the presentation
                ISlide slide3 = pres.Slides.AddClone(pres.Slides[0]);
                // Get the first effect of the first slide
                seq = slide3.Timeline.MainSequence;
                // Change the After animation effect to "Hide After Animation" 
                foreach (IEffect effect in seq)
                    effect.AfterAnimationType = AfterAnimationType.HideAfterAnimation;
                
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}