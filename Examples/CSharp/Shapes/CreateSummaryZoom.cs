using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This sample demonstrates how to create a summary zoom using Aspose.Slides for .NET
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class CreateSummaryZoom
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "SummaryZoomPresentation.pptx");

            using (Presentation pres = new Presentation())
            {
                //Adds a new slide to the presentation
                ISlide slide = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
                slide.Background.FillFormat.FillType = FillType.Solid;
                slide.Background.FillFormat.SolidFillColor.Color = Color.Brown;
                slide.Background.Type = BackgroundType.OwnBackground;

                // Adds a new section to the presentation
                pres.Sections.AddSection("Section 1", slide);

                //Adds a new slide to the presentation
                slide = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
                slide.Background.FillFormat.FillType = FillType.Solid;
                slide.Background.FillFormat.SolidFillColor.Color = Color.Aqua;
                slide.Background.Type = BackgroundType.OwnBackground;

                // Adds a new section to the presentation
                pres.Sections.AddSection("Section 2", slide);

                //Adds a new slide to the presentation
                slide = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
                slide.Background.FillFormat.FillType = FillType.Solid;
                slide.Background.FillFormat.SolidFillColor.Color = Color.Chartreuse;
                slide.Background.Type = BackgroundType.OwnBackground;

                // Adds a new section to the presentation
                pres.Sections.AddSection("Section 3", slide);

                //Adds a new slide to the presentation
                slide = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
                slide.Background.FillFormat.FillType = FillType.Solid;
                slide.Background.FillFormat.SolidFillColor.Color = Color.DarkGreen;
                slide.Background.Type = BackgroundType.OwnBackground;

                // Adds a new section to the presentation
                pres.Sections.AddSection("Section 4", slide);

                // Adds a SummaryZoomFrame object
                ISummaryZoomFrame summaryZoomFrame = pres.Slides[0].Shapes.AddSummaryZoomFrame(150, 50, 300, 200);

                // Saves the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
