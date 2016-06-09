using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.ProgrammersGuide.Shapes
{
    class ApplyBevelEffects
    {
        public static void Run()
        {
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
            pres.Save(dataDir + "Bavel.pptx", SaveFormat.Pptx);
        }
    }
}
