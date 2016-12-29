using System.IO;
using Aspose.Slides;
using Aspose.Slides.Effects;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class ShadowEffects
    {
        public static void Run()
        {
            // ExStart:ShadowEffects
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate a PPTX class
            using (Presentation pres = new Presentation())
            {

                // Get reference of the slide
                ISlide sld = pres.Slides[0];

                // Add an AutoShape of Rectangle type
                IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50);


                // Add TextFrame to the Rectangle
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
                pres.Save(dataDir + "pres_out.pptx", SaveFormat.Pptx);
            }
            // ExEnd:ShadowEffects
        }
    }
}