using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Animation;
using Aspose.Slides.Export;

/*
The following example demonstrates how to use ObjectCenter and SlideCenter subtype for FadedZoom effect.
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Transitions
{
    public class AnimationFadedZoomSubtype
    {
        public static void Run()
        {
            // Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation())
            {
                // Create shapes for demonstration
                var shp1 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 50, 50);
                var shp2 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 0, 50, 50);

                // Add FadedZoom effects
                var ef1 = pres.Slides[0].Timeline.MainSequence.AddEffect(shp1, EffectType.FadedZoom, EffectSubtype.ObjectCenter, EffectTriggerType.OnClick);
                var ef2 = pres.Slides[0].Timeline.MainSequence.AddEffect(shp2, EffectType.FadedZoom, EffectSubtype.SlideCenter, EffectTriggerType.OnClick);

                // Save presentation
                pres.Save(RunExamples.OutPath + "AnimationFadedZoom-out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}
