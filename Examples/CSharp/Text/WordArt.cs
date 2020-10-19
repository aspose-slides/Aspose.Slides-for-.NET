using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using DataTable = System.Data.DataTable;

namespace CSharp.Presentations.Conversion
{
    // This example demonstrates creating 3D shape and appliing 3D effects to the text in it.

    public class WordArt
    {
        public static void Run()
        {
            string resultPath = Path.Combine(RunExamples.OutPath, "WordArt_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create shape and text frame
                IAutoShape shape = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 314, 122, 400, 215.433f);

                ITextFrame textFrame = shape.TextFrame;

                Portion portion = (Portion)textFrame.Paragraphs[0].Portions[0];
                portion.Text = "Aspose.Slides";
                FontData fontData = new FontData("Arial Black");
                portion.PortionFormat.LatinFont = fontData;
                portion.PortionFormat.FontHeight = 36;

                // Set format of the text
                portion.PortionFormat.FillFormat.FillType = FillType.Pattern;
                portion.PortionFormat.FillFormat.PatternFormat.ForeColor.Color = Color.DarkOrange;
                portion.PortionFormat.FillFormat.PatternFormat.BackColor.Color = Color.White;
                portion.PortionFormat.FillFormat.PatternFormat.PatternStyle = PatternStyle.SmallGrid;

                portion.PortionFormat.LineFormat.FillFormat.FillType = FillType.Solid;
                portion.PortionFormat.LineFormat.FillFormat.SolidFillColor.Color = Color.Black;

                // Add a shadow effect for the text
                portion.PortionFormat.EffectFormat.EnableOuterShadowEffect();
                portion.PortionFormat.EffectFormat.OuterShadowEffect.ShadowColor.Color = Color.Black;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.ScaleHorizontal = 100;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.ScaleVertical = 65;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.BlurRadius = 4.73;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.Direction = 230;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.Distance = 2;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.SkewHorizontal = 30;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.SkewVertical = 0;
                portion.PortionFormat.EffectFormat.OuterShadowEffect.ShadowColor.ColorTransform.Add(ColorTransformOperation.SetAlpha, 0.32f);

                // Add reflection
                portion.PortionFormat.EffectFormat.EnableReflectionEffect();
                portion.PortionFormat.EffectFormat.ReflectionEffect.BlurRadius = 0.5;
                portion.PortionFormat.EffectFormat.ReflectionEffect.Distance = 4.72;
                portion.PortionFormat.EffectFormat.ReflectionEffect.StartPosAlpha = 0f;
                portion.PortionFormat.EffectFormat.ReflectionEffect.EndPosAlpha = 60f;
                portion.PortionFormat.EffectFormat.ReflectionEffect.Direction = 90;
                portion.PortionFormat.EffectFormat.ReflectionEffect.ScaleHorizontal = 100;
                portion.PortionFormat.EffectFormat.ReflectionEffect.ScaleVertical = -100;
                portion.PortionFormat.EffectFormat.ReflectionEffect.StartReflectionOpacity = 60f;
                portion.PortionFormat.EffectFormat.ReflectionEffect.EndReflectionOpacity = 0.9f;
                portion.PortionFormat.EffectFormat.ReflectionEffect.RectangleAlign = RectangleAlignment.BottomLeft;

                // Add glow effect
                portion.PortionFormat.EffectFormat.EnableGlowEffect();
                portion.PortionFormat.EffectFormat.GlowEffect.Color.R = 255;
                portion.PortionFormat.EffectFormat.GlowEffect.Color.ColorTransform.Add(ColorTransformOperation.SetAlpha, 0.54f);
                portion.PortionFormat.EffectFormat.GlowEffect.Radius = 7;

                // Add transformation
                textFrame.TextFrameFormat.Transform = TextShapeType.ArchUpPour;

                // Add 3D effects to the shape
                shape.ThreeDFormat.BevelBottom.BevelType = BevelPresetType.Circle;
                shape.ThreeDFormat.BevelBottom.Height = 10.5;
                shape.ThreeDFormat.BevelBottom.Width = 10.5;

                shape.ThreeDFormat.BevelTop.BevelType = BevelPresetType.Circle;
                shape.ThreeDFormat.BevelTop.Height = 12.5;
                shape.ThreeDFormat.BevelTop.Width = 11;

                shape.ThreeDFormat.ExtrusionColor.Color = Color.Orange;
                shape.ThreeDFormat.ExtrusionHeight = 6;

                shape.ThreeDFormat.ContourColor.Color = Color.DarkRed;
                shape.ThreeDFormat.ContourWidth = 1.5;

                shape.ThreeDFormat.Depth = 3;

                shape.ThreeDFormat.Material = MaterialPresetType.Plastic;

                shape.ThreeDFormat.LightRig.Direction = LightingDirection.Top;
                shape.ThreeDFormat.LightRig.LightType = LightRigPresetType.Balanced;
                shape.ThreeDFormat.LightRig.SetRotation(0, 0, 40);

                shape.ThreeDFormat.Camera.CameraType = CameraPresetType.PerspectiveContrastingRightFacing;

                // Add 3D effects to the text
                textFrame = shape.TextFrame;

                textFrame.TextFrameFormat.ThreeDFormat.BevelBottom.BevelType = BevelPresetType.Circle;
                textFrame.TextFrameFormat.ThreeDFormat.BevelBottom.Height = 3.5;
                textFrame.TextFrameFormat.ThreeDFormat.BevelBottom.Width = 3.5;

                textFrame.TextFrameFormat.ThreeDFormat.BevelTop.BevelType = BevelPresetType.Circle;
                textFrame.TextFrameFormat.ThreeDFormat.BevelTop.Height = 12.5;
                textFrame.TextFrameFormat.ThreeDFormat.BevelTop.Width = 11;

                textFrame.TextFrameFormat.ThreeDFormat.ExtrusionColor.Color = Color.Orange;
                textFrame.TextFrameFormat.ThreeDFormat.ExtrusionHeight = 6;

                textFrame.TextFrameFormat.ThreeDFormat.ContourColor.Color = Color.DarkRed;
                textFrame.TextFrameFormat.ThreeDFormat.ContourWidth = 1.5;

                textFrame.TextFrameFormat.ThreeDFormat.Depth = 3;

                textFrame.TextFrameFormat.ThreeDFormat.Material = MaterialPresetType.Plastic;

                textFrame.TextFrameFormat.ThreeDFormat.LightRig.Direction = LightingDirection.Top;
                textFrame.TextFrameFormat.ThreeDFormat.LightRig.LightType = LightRigPresetType.Balanced;
                textFrame.TextFrameFormat.ThreeDFormat.LightRig.SetRotation(0, 0, 40);

                textFrame.TextFrameFormat.ThreeDFormat.Camera.CameraType = CameraPresetType.PerspectiveContrastingRightFacing;

                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
