using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells.Drawing;
using Aspose.Slides.Util;

/*
The example demonstrates the use of the SlideUtil.FindShapesByPlaceholderType 
and SlideUtil.GetTextBoxesContainsText methods to find text on a slide.
*/

namespace CSharp.Text
{
    class GetPlaceholderTextExample
    {
        public static void Run()
        {
            using (var presentation = new Presentation())
            {
                // Add new slide based on LayoutSlides[0]
                ISlide slide = presentation.Slides.AddEmptySlide(presentation.LayoutSlides[0]);

                // Search for specified text in a slide, including its layout (layout template text)
                foreach (ITextFrame textFrame in SlideUtil.GetTextBoxesContainsText(slide, "Click", true))
                {
                    // Set text for TextFrame
                    Console.WriteLine("A text block with the specified text was found.");
                }

                // Find all “Text” placeholders on a slide:
                foreach (IShape shape in SlideUtil.FindShapesByPlaceholderType(slide, PlaceholderType.CenteredTitle))
                {
                    Console.WriteLine("Placeholder of type PlaceholderType.CenteredTitle was found.");
                }
            }
        }
    }
}