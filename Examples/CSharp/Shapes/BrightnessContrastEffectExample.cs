using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;

/*
The following sample code shows how to use IBrightnessContrast, IBrightnessContrastEffectiveData interfaces 
and AddBrightnessContrastEffect method to get brightness and contrast values of BrightnessContrast effect if they are present.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class BrightnessContrastEffectExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Shapes(), "BrightnessContrast.pptx");

            using (var presentation = new Presentation(presentationName))
            {
                // Get slide
                ISlide slide = presentation.Slides[0];

                // Get picture frame
                IPictureFrame pictureFrame = (IPictureFrame)(slide.Shapes[0]);

                // Get image transform operations
                IImageTransformOperationCollection imageTransform = pictureFrame.PictureFormat.Picture.ImageTransform;
                foreach (var effect in imageTransform)
                {
                    if (effect is IBrightnessContrast)
                    {
                        // Get brightness and contrast values
                        IBrightnessContrast brightnessContrast = (IBrightnessContrast)effect;
                        IBrightnessContrastEffectiveData brightnessContrastData = brightnessContrast.GetEffective();

                        Console.WriteLine("Brightness value = {0}", brightnessContrastData.Brightness);
                        Console.WriteLine("Contrast value = {0}", brightnessContrastData.Contrast);
                    }
                }
            }
        }
    }
}
