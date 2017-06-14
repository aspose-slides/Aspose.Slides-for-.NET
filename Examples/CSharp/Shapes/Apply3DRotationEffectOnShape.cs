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
    class Apply3DRotationEffecrOnShapes
    {
        public static void Run()
        {
            //ExStart:Apply3DRotationEffecrOnShapes
            // The path to the documents directory.                    
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create an instance of Presentation class
            Presentation pres = new Presentation();
            IShape autoShape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 30, 30, 200, 200);

            autoShape.ThreeDFormat.Depth = 6;
            autoShape.ThreeDFormat.Camera.SetRotation(40, 35, 20);
            autoShape.ThreeDFormat.Camera.CameraType = CameraPresetType.IsometricLeftUp;
            autoShape.ThreeDFormat.LightRig.LightType = LightRigPresetType.Balanced;

            autoShape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Line, 30, 300, 200, 200);
            autoShape.ThreeDFormat.Depth = 6;
            autoShape.ThreeDFormat.Camera.SetRotation(0, 35, 20);
            autoShape.ThreeDFormat.Camera.CameraType = CameraPresetType.IsometricLeftUp;
            autoShape.ThreeDFormat.LightRig.LightType = LightRigPresetType.Balanced;

          
            pres.Save(dataDir + "Rotation_out.pptx", SaveFormat.Pptx);
            //ExEnd:Apply3DRotationEffecrOnShapes
        }
    }
}
