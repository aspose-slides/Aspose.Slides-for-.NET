using Aspose.Slides.Export;
using Aspose.Slides;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Aspose.Slides.Examples.CSharp.Rendering.Printing
{
    class Rendering3D
    {

        // This example demonstrates creating and rendering presentation with 3D graphics.

        public static void Run()
        {
            string outPptxFile = Path.Combine(RunExamples.OutPath, "sandbox_3d.pptx");
            string outPngFile = Path.Combine(RunExamples.OutPath, "sample_3d.png");

            using (Presentation pres = new Presentation())
            {
                IAutoShape shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 200, 150, 200, 200);
                shape.TextFrame.Text = "3D";
                shape.TextFrame.Paragraphs[0].ParagraphFormat.DefaultPortionFormat.FontHeight = 64;

                shape.ThreeDFormat.Camera.CameraType = CameraPresetType.OrthographicFront;
                shape.ThreeDFormat.Camera.SetRotation(20, 30, 40);
                shape.ThreeDFormat.LightRig.LightType = LightRigPresetType.Flat;
                shape.ThreeDFormat.LightRig.Direction = LightingDirection.Top;
                shape.ThreeDFormat.Material = MaterialPresetType.Powder;
                shape.ThreeDFormat.ExtrusionHeight = 100;
                shape.ThreeDFormat.ExtrusionColor.Color = Color.Blue;

                pres.Slides[0].GetImage(2, 2).Save(outPngFile, ImageFormat.Png);
                pres.Save(outPptxFile, SaveFormat.Pptx);
            }
        }
    }
}


