using System;
using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;

/*
This code demonstrates iteration over all Presentation shapes and out to console if the shape is a text box or not (if the shape is AutoShape).
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class IsTextShape
    {
        public static void Run()
        {
            string presentationPath = RunExamples.GetDataDir_Shapes() + "CheckTextShapes.pptx";

            using (Presentation presentation = new Presentation(presentationPath))
            {
                foreach (IShape shape in presentation.Slides[0].Shapes)
                {
                    if (shape is AutoShape autoShape)
                    {
                        Console.WriteLine(autoShape.IsTextBox ? "shape is text box" : "shape is not text box");
                    }
                }
            }
        }
    }
}
