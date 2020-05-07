using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Util;
using Aspose.Slides.Export;
using Aspose.Slides.MathText;

/*
This example demonstrates of using API for creation a mathematical expression for Pythagorean theorem.
*/
namespace CSharp.Shapes
{
    class ShapesAlignment
    {
        public static void Run()
        {
            //Path for output presentation
            string outpptxFile = Path.Combine(RunExamples.OutPath, "ShapesAlignment_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create some shapes
                ISlide slide = pres.Slides[0];
                IAutoShape shape1 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
                IAutoShape shape2 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 250, 200, 100, 100);
                IAutoShape shape3 = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 400, 300, 100, 100);

                // Here we align two shapes using their indexes
                SlideUtil.AlignShapes(ShapesAlignmentType.AlignMiddle, true, slide, new int[]
                {
                    slide.Shapes.IndexOf(shape1),
                    slide.Shapes.IndexOf(shape2)
                });

                // Here we aling all shapes int the slide
                SlideUtil.AlignShapes(ShapesAlignmentType.AlignMiddle, true, pres.Slides[0].Shapes);

                pres.Save(outpptxFile, SaveFormat.Pptx);
            }
        }
    }
}
