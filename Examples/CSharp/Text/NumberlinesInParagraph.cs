using System;
using System.Drawing;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.Ink;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/*
This example shows how to get lines count in a paragraph.
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class NumberlinesInParagraph
    {
        public static void Run()
        {
            using (Presentation presentation = new Presentation())
            {
                ISlide sld = presentation.Slides[0];
                IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50);
                IParagraph para = ashp.TextFrame.Paragraphs[0];
                IPortion portion = para.Portions[0];
                portion.Text = "Aspose Paragraph GetLinesCount() Example";

                Console.WriteLine("Lines Count = {0}", para.GetLinesCount());

                // Change shape width
                ashp.Width = 250;
                Console.WriteLine("Lines Count after changing shape width = {0}", para.GetLinesCount());
            }
        }
    }
}
