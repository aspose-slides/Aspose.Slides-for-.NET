using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This example demonstrates how to create a zoom frame with different images 
and shows how to change the formatting of a zoom frame.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class CreateZoomFrame
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "ZoomFramePresentation.pptx");

            // Path to source image
            string imagePath = Path.Combine(RunExamples.GetDataDir_Shapes(), "aspose-logo.jpg");

            using (Presentation pres = new Presentation())
            {
                //Add new slides to the presentation
                ISlide slide2 = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);
                ISlide slide3 = pres.Slides.AddEmptySlide(pres.Slides[0].LayoutSlide);

                // Create a background for the second slide
                slide2.Background.Type = BackgroundType.OwnBackground;
                slide2.Background.FillFormat.FillType = FillType.Solid;
                slide2.Background.FillFormat.SolidFillColor.Color = Color.Cyan;

                // Create a text box for the second slide
                IAutoShape autoshape = slide2.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 500, 200);
                autoshape.TextFrame.Text = "Second Slide";

                // Create a background for the third slide
                slide3.Background.Type = BackgroundType.OwnBackground;
                slide3.Background.FillFormat.FillType = FillType.Solid;
                slide3.Background.FillFormat.SolidFillColor.Color = Color.DarkKhaki;

                // Create a text box for the third slide
                autoshape = slide3.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 500, 200);
                autoshape.TextFrame.Text = "Trird Slide";

                // Add ZoomFrame objects with slide preview
                var zoomFrame1 = pres.Slides[0].Shapes.AddZoomFrame(20, 20, 250, 200, slide2);

                // Add ZoomFrame objects with custom image
                // Create a new image for the zoom object
                IPPImage image = pres.Images.AddImage(Images.FromFile(imagePath));
                var zoomFrame2 = pres.Slides[0].Shapes.AddZoomFrame(200, 250, 250, 100, slide3, image);

                // Set a zoom frame format for the zoomFrame2 object
                zoomFrame2.LineFormat.Width = 5;
                zoomFrame2.LineFormat.FillFormat.FillType = FillType.Solid;
                zoomFrame2.LineFormat.FillFormat.SolidFillColor.Color = Color.HotPink;
                zoomFrame2.LineFormat.DashStyle = LineDashStyle.DashDot;

                // Do not show background for zoomFrame1 object
                zoomFrame1.ShowBackground = false;


                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
