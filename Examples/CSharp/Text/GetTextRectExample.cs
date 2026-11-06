using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to get text rectangle defined by the parent shape’s geometry.
*/

namespace CSharp.Text
{
    class GetTextRectExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "CustomShapes.pptx");

            using (Presentation presentation = new Presentation(presentationName))
            {
                IAutoShape shape = (IAutoShape)presentation.Slides[0].Shapes[0];
                RectangleF textRectangle = shape.TextFrame.GetTextRect();

                 Console.WriteLine(
                    "Shape rect: X={0}, Y={1}, Width={2}, Height={3}",
                    shape.X,
                    shape.Y,
                    shape.Width,
                    shape.Height);
               Console.WriteLine(
                    "Text rect: X={0}, Y={1}, Width={2}, Height={3}",
                    textRectangle.X,
                    textRectangle.Y,
                    textRectangle.Width,
                    textRectangle.Height);
            }
        }
    }
}