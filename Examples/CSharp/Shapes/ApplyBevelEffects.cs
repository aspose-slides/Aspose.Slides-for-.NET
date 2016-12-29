using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class ApplyBevelEffects
    {
        public static void Run()
        {
            //ExStart:ApplyBevelEffects
            // The path to the documents directory.                    
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create an instance of Presentation class
            Presentation pres = new Presentation();
            ISlide slide = pres.Slides[0];

            // Add a shape on slide
            IAutoShape shape = slide.Shapes.AddAutoShape(ShapeType.Ellipse, 30, 30, 100, 100);
            shape.FillFormat.FillType = FillType.Solid;
            shape.FillFormat.SolidFillColor.Color = Color.Green;
            ILineFillFormat format = shape.LineFormat.FillFormat;
            format.FillType = FillType.Solid;
            format.SolidFillColor.Color = Color.Orange;
            shape.LineFormat.Width = 2.0;

            // Set ThreeDFormat properties of shape
            shape.ThreeDFormat.Depth = 4;
            shape.ThreeDFormat.BevelTop.BevelType = BevelPresetType.Circle;
            shape.ThreeDFormat.BevelTop.Height = 6;
            shape.ThreeDFormat.BevelTop.Width = 6;
            shape.ThreeDFormat.Camera.CameraType = CameraPresetType.OrthographicFront;
            shape.ThreeDFormat.LightRig.LightType = LightRigPresetType.ThreePt;
            shape.ThreeDFormat.LightRig.Direction = LightingDirection.Top;

            // Write the presentation as a PPTX file
            pres.Save(dataDir + "Bavel_out.pptx", SaveFormat.Pptx);
            //ExEnd:ApplyBevelEffects
        }
    }
}
