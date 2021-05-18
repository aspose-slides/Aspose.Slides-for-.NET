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
                // Create slides array
                for (int slideNumber = 0; slideNumber < 5; slideNumber++)
                {
                    //Add new slides to presentation
                    ISlide slide = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);

                    // Create a background for the slide
                    slide.Background.Type = BackgroundType.OwnBackground;
                    slide.Background.FillFormat.FillType = FillType.Solid;
                    slide.Background.FillFormat.SolidFillColor.Color = Color.DarkKhaki;

                    // Create a text box for the slide
                    IAutoShape autoshape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 500, 200);
                    autoshape.TextFrame.Text = String.Format("Slide - {0}", slideNumber + 2);
                }

                // Create zoom objects for all slides in the first slide
                for (int slideNumber = 1; slideNumber < pres.Slides.Count; slideNumber++)
                {
                    int x = (slideNumber - 1) * 100;
                    int y = (slideNumber - 1) * 100;
                    IZoomFrame zoomFrame = pres.Slides[0].Shapes.AddZoomFrame(x, y, 150, 120, pres.Slides[slideNumber]);

                    // Set the ReturnToParent property to return to the first slide
                    zoomFrame.ReturnToParent = true;
                }

                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
