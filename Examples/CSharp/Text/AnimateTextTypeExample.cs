using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example shows how to set the animate text type of an animation effect to "By Letter" value. 
Please note that the text animation type allows you to set the following text animation types:
- animate all text at once
- animate text by word
 -animate text by letter
*/

namespace CSharp.Text
{
    class AnimateTextTypeExample
    {
        public static void Run()
        {
            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "AnimateTextEffect_out.pptx");

            using (Presentation presentation = new Presentation())
            {
                IAutoShape oval = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Ellipse, 100, 100, 300, 150);
                oval.TextFrame.Text = "The new animated text";

                // Get anomation timeline.
                IAnimationTimeLine timeline = presentation.Slides[0].Timeline;

                // Set the effect of the first slide.
                IEffect effect = timeline.MainSequence.AddEffect(oval, EffectType.Appear, EffectSubtype.None, EffectTriggerType.OnClick);

                // Set the effect Animate text type to "By letter".
                effect.AnimateTextType = AnimateTextType.ByLetter;

                // Set the delay between animated text parts.
                effect.DelayBetweenTextParts = -1.5f;

                // Save presentation.
                presentation.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}

