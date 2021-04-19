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
This example demonstrates of using SlideUtil.AlignShapes method.
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
                ISlide slide = pres.Slides[0];
                // Create some shapes
                slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 100, 100);
                slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 100, 100);
                slide.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 300, 100, 100);
                // Aligning all shapes within IBaseSlide.
                SlideUtil.AlignShapes(ShapesAlignmentType.AlignBottom, true, pres.Slides[0]);

                slide = pres.Slides.AddEmptySlide(slide.LayoutSlide);
                // Add group shape
                IGroupShape groupShape = slide.Shapes.AddGroupShape();
                // Create some shapes to the group shape
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 350, 50, 50, 50);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 450, 150, 50, 50);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 550, 250, 50, 50);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 650, 350, 50, 50);
                // Aligning all shapes within IGroupShape.
                SlideUtil.AlignShapes(ShapesAlignmentType.AlignLeft, false, groupShape);

                slide = pres.Slides.AddEmptySlide(slide.LayoutSlide);
                // Add group shape
                groupShape = slide.Shapes.AddGroupShape();
                // Create some shapes to the group shape
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 350, 50, 50, 50);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 450, 150, 50, 50);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 550, 250, 50, 50);
                groupShape.Shapes.AddAutoShape(ShapeType.Rectangle, 650, 350, 50, 50);
                // Aligning shapes with specified indexes within IGroupShape.
                SlideUtil.AlignShapes(ShapesAlignmentType.AlignLeft, false, groupShape, new int[] { 0, 2 });

                // Save presentation
                pres.Save(outpptxFile, SaveFormat.Pptx);
            }
        }
    }
}
