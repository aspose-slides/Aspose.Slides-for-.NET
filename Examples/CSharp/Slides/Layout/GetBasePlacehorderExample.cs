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
This example shows how to get all (master/layout/slide) animated effects of a placeholder shape using the Shape.GetBasePlaceholder method.
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Layout
{
    class GetBasePlacehorderExample
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Layout(), "placeholder.pptx");

            using (Presentation presentation = new Presentation(presentationName))
            {
                ISlide slide = presentation.Slides[0];
                IShape shape = slide.Shapes[0];
                IEffect[] shapeEffects = slide.LayoutSlide.Timeline.MainSequence.GetEffectsByShape(shape);
                Console.WriteLine("Shape effects count = {0}", shapeEffects.Length);

                IShape layoutShape = shape.GetBasePlaceholder();
                IEffect[] layoutShapeEffects = slide.LayoutSlide.Timeline.MainSequence.GetEffectsByShape(layoutShape);
                Console.WriteLine("Layout shape effects count = {0}", layoutShapeEffects.Length);

                IShape masterShape = layoutShape.GetBasePlaceholder();
                IEffect[] masterShapeEffects = slide.LayoutSlide.MasterSlide.Timeline.MainSequence.GetEffectsByShape(masterShape);
                Console.WriteLine("Master shape effects count = {0}", masterShapeEffects.Length);
            }
        }
    }
}
