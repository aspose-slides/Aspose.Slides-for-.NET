using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This example demonstrates how to retrieve the fill color defined by the presentation’s color scheme.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class SolidFillSchemeColorExamole
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationPath = Path.Combine(RunExamples.GetDataDir_Shapes(), "FillColor.pptx");

            using (var presentation = new Presentation(presentationPath))
            {
                foreach (var shape in presentation.Slides[0].Shapes)
                {
                    if (shape is AutoShape)
                    {
                        var fillFormat = ((AutoShape)shape).TextFrame.Paragraphs[0].Portions[0].PortionFormat.FillFormat.GetEffective();
                        Console.WriteLine("Fill color: " + fillFormat.SolidFillColor);
                        Console.WriteLine("Fill scheme color: " + fillFormat.SolidFillSchemeColor);
                    }
                }
            }
        }
    }
}