using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This sample demonstrates how to delete cropped areas of the fill Picture 
(This can help to reduce the size of the presentation).
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class DeletePictureCroppedAreasExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Shapes(), "CroppedImage.pptx");
            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "CroppedImage-out.pptx");


            using (Presentation pres = new Presentation(presentationName))
            {
                // Gets the first slide
                ISlide slide = pres.Slides[0];

                // Gets the PictureFrame
                IPictureFrame picFrame = slide.Shapes[0] as IPictureFrame;

                // Deletes cropped areas of the PictureFrame image
                IPPImage croppedImage = picFrame.PictureFormat.DeletePictureCroppedAreas();

                // Save result
                pres.Save(outFilePath, SaveFormat.Pptx);

                // Check size
                Console.WriteLine("Source presentation length\t = {0}", new FileInfo(presentationName).Length);
                Console.WriteLine("Resulting presentation length\t = {0}", new FileInfo(outFilePath).Length);
            }
        }
    }
}
