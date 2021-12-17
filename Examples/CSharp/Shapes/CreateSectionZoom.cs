using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This sample demonstrates how to create a section zoom using Aspose.Slides for .NET
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class CreateSectionZoom
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "SectionZoomPresentation.pptx");

            using (Presentation pres = new Presentation())
            {
                //Adds a new slide to the presentation
                ISlide slide = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
                slide.Background.FillFormat.FillType = FillType.Solid;
                slide.Background.FillFormat.SolidFillColor.Color = Color.YellowGreen;
                slide.Background.Type = BackgroundType.OwnBackground;

                // Adds a new Section to the presentation
                pres.Sections.AddSection("Section 1", slide);

                // Adds a SectionZoomFrame object
                ISectionZoomFrame sectionZoomFrame = pres.Slides[0].Shapes.AddSectionZoomFrame(20, 20, 300, 200, pres.Sections[1]);

                // Saves the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
