//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Effects;
using Aspose.Slides.Export;

namespace CSharp.Text
{
    public class ShadowEffects
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate a PPTX class
            using (Presentation pres = new Presentation())
            {

                //Get reference of the slide
                ISlide sld = pres.Slides[0];

                //Add an AutoShape of Rectangle type
                IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50);


                //Add TextFrame to the Rectangle
                ashp.AddTextFrame("Aspose TextBox");

                // Disable shape fill in case we want to get shadow of text
                ashp.FillFormat.FillType = FillType.NoFill;

                // Add outer shadow and set all necessary parameters
                ashp.EffectFormat.EnableOuterShadowEffect();
                IOuterShadow shadow = ashp.EffectFormat.OuterShadowEffect;
                shadow.BlurRadius = 4.0;
                shadow.Direction = 45;
                shadow.Distance = 3;
                shadow.RectangleAlign = RectangleAlignment.TopLeft;
                shadow.ShadowColor.PresetColor = PresetColor.Black;

                //Write the presentation to disk
                pres.Save(dataDir + "pres.pptx", SaveFormat.Pptx);
            }

        }
    }
}