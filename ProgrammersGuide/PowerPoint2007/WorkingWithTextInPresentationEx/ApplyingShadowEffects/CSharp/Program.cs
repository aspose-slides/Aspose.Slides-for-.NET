//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;
using Aspose.Slides.Effects;

namespace ApplyingShadowEffects
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate a PPTX class
            PresentationEx pres = new PresentationEx();


            //Get first slide
            SlideEx sld = pres.Slides[0];

            //Add an AutoShape of Rectangle type
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 75, 150, 50);
            AutoShapeEx ashp = (AutoShapeEx)sld.Shapes[idx];

            //Add TextFrame to the Rectangle
            ashp.AddTextFrame("Aspose TextBox");

            // Disable shape fill in case we want to get shadow of text.
            ashp.FillFormat.FillType = FillTypeEx.NoFill;

            // Add outer shadow and set all necessary parameters
            OuterShadow shadow = new OuterShadow();
            ashp.EffectFormat.OuterShadowEffect = shadow;
            shadow.BlurRadius = 4.0;
            shadow.Direction = 45;
            shadow.Distance = 3;
            shadow.RectangleAlign = RectangleAlignmentEx.TopLeft;
            shadow.ShadowColor.PresetColor = PresetColorEx.Black;

            //Write the presentation to disk
            pres.Write(dataDir + "output.pptx");


        }
    }
}