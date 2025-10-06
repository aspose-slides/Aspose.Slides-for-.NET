using System;
using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;

/*
This code demonstrates an operation where we added a picture for a slide background, added Duotone effect with styled colors, 
and then we got the effective duotone colors with which the background will be rendered.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class DuotoneEffectsPVI
    {
        public static void Run()
        {
            using (Presentation presentation = new Presentation())
            {
                string imagePath = RunExamples.GetDataDir_Shapes() + "aspose-logo.jpg";

                // Add image to presentation
                IPPImage backgroundImage = presentation.Images.AddImage(Images.FromFile(imagePath));

                // Set background in first slide
                presentation.Slides[0].Background.Type = BackgroundType.OwnBackground;
                presentation.Slides[0].Background.FillFormat.FillType = FillType.Picture;
                presentation.Slides[0].Background.FillFormat.PictureFillFormat.Picture.Image = backgroundImage;

                // Add Duotone effect to background
                IDuotone duotone = presentation.Slides[0].Background.FillFormat.PictureFillFormat.Picture.ImageTransform
                    .AddDuotoneEffect();

                // Set Doutone properties
                duotone.Color1.ColorType = ColorType.Scheme;
                duotone.Color1.SchemeColor = SchemeColor.Accent1;
                duotone.Color2.ColorType = ColorType.Scheme;
                duotone.Color2.SchemeColor = SchemeColor.Dark2;

                // Get Effective values of the Duotone effect
                IDuotoneEffectiveData duotoneEffective = duotone.GetEffective();

                // Show effective values
                Console.WriteLine("Duotone effective color1: " + duotoneEffective.Color1);
                Console.WriteLine("Duotone effective color2: " + duotoneEffective.Color2);
            }
        }
    }
}
