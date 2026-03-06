using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This example demonstrates how to get visual bounds of shape that represents the final rendered appearance and may extend 
beyond the shape’s original geometry. The returned bounds take into account rendering-time factors such as rotation, 
stroke width, text overflow, SmartArt layout, and grouping.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class ShapeVisualBoundsExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationPath = Path.Combine(RunExamples.GetDataDir_Shapes(), "Shapes.pptx");

            using (var presentation = new Presentation(presentationPath))
            {
                Shape shape = (Shape)presentation.Slides[0].Shapes[0];

                RectangleF visualBounds = shape.GetVisualBounds();

                Console.WriteLine(
                    $"Visual bounds: X={visualBounds.X}, Y={visualBounds.Y}, " +
                    $"Width={visualBounds.Width}, Height={visualBounds.Height}");
            }
        }
    }
}
