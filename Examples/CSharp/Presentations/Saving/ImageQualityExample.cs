using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following code sample demonstrates how to save the first image in the presentation collection as a JPEG with varying quality.
*/

namespace CSharp.Presentations.Saving
{
    class ImageQualityExample
    {
        public static void Run()
        {
            //Path for source presentation
            string pptxFile = Path.Combine(RunExamples.GetDataDir_PresentationSaving(), "ImageQuality.pptx");
            //Out path
            string imagePath = Path.Combine(RunExamples.OutPath, "ImageQuality-out.jpg");

            using (Presentation presentation = new Presentation(pptxFile))
            {
                var image = presentation.Images[0].Image;

                // Saves the first image to the memory stream in JPEG format with quality 80.
                using (MemoryStream ms = new MemoryStream())
                {
                    image.Save(ms, ImageFormat.Jpeg, 100);
                }

                // Saves the first image to the file in JPEG format with high quality.
                image.Save(imagePath, ImageFormat.Jpeg, 100);
            }
        }
    }
}
