using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
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
The example below demonstrates how to set sketchy type for a shape.
Please pay attention that not all versions of PowerPoint can display sketched shapes.
*/
namespace CSharp.Shapes
{
    class SketchedShapes
    {
        public static void Run()
        {
            //Path for output presentation
            string outPptxFile = Path.Combine(RunExamples.OutPath, "SketchedShapes_out.pptx");
            string outPngFile = Path.Combine(RunExamples.OutPath, "SketchedShapes_out.png");

            using (Presentation pres = new Presentation())
            {
                IAutoShape shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 20, 20, 300, 150);
                shape.FillFormat.FillType = FillType.NoFill;

                // Transform shape to sketch of a freehand style
                shape.LineFormat.SketchFormat.SketchType = LineSketchType.Scribble;

                pres.Slides[0].GetImage(4/3f, 4/3f).Save(outPngFile, Aspose.Slides.ImageFormat.Png);
                pres.Save(outPptxFile, SaveFormat.Pptx);
            }
        }
    }
}
