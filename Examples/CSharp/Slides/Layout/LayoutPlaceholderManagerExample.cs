using System.IO;
using Aspose.Slides.Export;

/*
The following example shows how to add new placeholder shapes to the Layout slide.
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Layout
{
    class LayoutPlaceholderManagerExample
    {
        public static void Run()
        {
            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "placeholders.pptx");

            using (var pres = new Presentation())
            {
                // Getting the Blank layout slide.
                ILayoutSlide layout = pres.LayoutSlides.GetByType(SlideLayoutType.Blank);

                // Getting the placeholder manager of the layout slide.
                ILayoutPlaceholderManager placeholderManager = layout.PlaceholderManager;

                // Adding different placeholders to the Blank layout slide.
                placeholderManager.AddContentPlaceholder(10, 10, 300, 200);
                placeholderManager.AddVerticalTextPlaceholder(350, 10, 200, 300);
                placeholderManager.AddChartPlaceholder(10, 350, 300, 300);
                placeholderManager.AddTablePlaceholder(350, 350, 300, 200);

                // Adding the new slide with Blank layout.
                ISlide newSlide = pres.Slides.AddEmptySlide(layout);

                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
