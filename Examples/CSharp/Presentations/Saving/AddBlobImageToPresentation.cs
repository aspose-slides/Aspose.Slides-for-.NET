using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Saving
{
    class AddBlobImageToPresentation
    {
        public static void Run() {

            //ExStart:AddBlobImageToPresentation

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();

            string pathToLargeImage = dataDir + "large_image.jpg";

            // create a new presentation which will contain this image
            using (Presentation pres = new Presentation())
            {
                using (FileStream fileStream = new FileStream(pathToLargeImage, FileMode.Open))
                {
                    // let's add the image to the presentation - we choose KeepLocked behavior, because we not
                    // have an intent to access the "largeImage.png" file.
                    IPPImage img = pres.Images.AddImage(fileStream, LoadingStreamBehavior.KeepLocked);
                    pres.Slides[0].Shapes.AddPictureFrame(ShapeType.Rectangle, 0, 0, 300, 200, img);

                    // save the presentation. Despite that the output presentation will be
                    // large, the memory consumption will be low the whole lifetime of the pres object
                    pres.Save(dataDir + "presentationWithLargeImage.pptx", SaveFormat.Pptx);
                }
            }

            //ExEnd:AddBlobImageToPresentation

        }
    }
}
