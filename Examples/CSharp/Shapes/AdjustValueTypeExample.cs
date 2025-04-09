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
 This example demonstrates how to use types of Adjustment points.
*/
namespace CSharp.Shapes
{
    class AdjustValueTypeExample
    {
        public static void Run()
        {
            //Path for presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Shapes(), "PresetGeometry.pptx");

            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "PresetGeometry_out.pptx");

            using (var pres = new Presentation(presentationName))
            {
                var shape = (IAutoShape)pres.Slides[0].Shapes[0];

                // Show all adjustment point and its types for a RoundRectangle
                Console.WriteLine("Adjustment types for a Rectangle:");
                for (int i=0 ; i < shape.Adjustments.Count ; i++)
                {
                    Console.WriteLine("\tType for point {0} is \"{1}\"", i, shape.Adjustments[i].Type);
                }
                // Change value of an adjustment point
                if (shape.Adjustments[0].Type == ShapeAdjustmentType.CornerSize)
                {
                    shape.Adjustments[0].AngleValue *= 2;
                }

                // Show all adjustment point and its types for an RightArrow
                var shape1 = (IAutoShape)pres.Slides[0].Shapes[1];
                Console.WriteLine("Adjustment types for an Arrow:");
                for (int i = 0; i < shape1.Adjustments.Count; i++)
                {
                    Console.WriteLine("\tType for point {0} is \"{1}\"", i, shape1.Adjustments[i].Type);
                }
                // Change value of adjustment points
                if (shape1.Adjustments[0].Type == ShapeAdjustmentType.ArrowTailThickness)
                {
                    shape1.Adjustments[0].AngleValue /= 3;
                }
                if (shape1.Adjustments[1].Type == ShapeAdjustmentType.ArrowheadLength)
                {
                    shape1.Adjustments[1].AngleValue /= 2;
                }

                // Save the presentation
                pres.Save(outFilePath, SaveFormat.Pptx);

            }
        }
    }
}
