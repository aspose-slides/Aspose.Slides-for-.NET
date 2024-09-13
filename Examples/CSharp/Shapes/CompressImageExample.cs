using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example shows how to compresses an image by reducing its size based on the shape size and specified resolution, with the option to delete cropped areas.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class CompressImageExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Shapes(), "CroppedImage.pptx");
            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "CompressImage-out.pptx");


            using (Presentation pres = new Presentation(presentationName))
            {
                ISlide slide = pres.Slides[0];

                // Get the PictureFrame from the slide
                IPictureFrame picFrame = slide.Shapes[0] as IPictureFrame;

                // Compress the image with a target resolution of 150 DPI (Web resolution) and remove cropped areas
                bool result = picFrame.PictureFormat.CompressImage(true, 150f);

                // Check the result of the compression
                if (result)
                {
                    Console.WriteLine("Image successfully compressed.");
                }
                else
                {
                    Console.WriteLine("Image compression failed or no changes were necessary.");
                }

                // Save result
                pres.Save(outFilePath, SaveFormat.Pptx);

                // Check size
                Console.WriteLine("Source presentation length\t = {0}", new FileInfo(presentationName).Length);
                Console.WriteLine("Resulting presentation length\t = {0}", new FileInfo(outFilePath).Length);
            }
        }
    }
}
