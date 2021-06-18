using System;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.DOM.Ole;
using Aspose.Slides.Export;

/*
This sample demonstrates the output of information for all animated shapes in the main sequence for all slides in a presentation.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AnimationTargetShapes
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationFileName = Path.Combine(RunExamples.GetDataDir_Shapes(), "AnimationShapesExample.pptx");

            using (Presentation pres = new Presentation(presentationFileName))
            {
                foreach (ISlide slide in pres.Slides)
                {
                    foreach (IEffect effect in slide.Timeline.MainSequence)
                    {
                        Console.WriteLine(effect.Type + " animation effect is set to shape#" +
                                          effect.TargetShape.UniqueId +
                                          " on slide#" + slide.SlideNumber);
                    }
                }
            }
        }
    }
}