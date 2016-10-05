using System.IO;
using Aspose.Slides;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class ThumbnailFromSlide
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            // Instantiate a Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "ThumbnailFromSlide.pptx"))
            {

                //Access the first slide
                ISlide sld = pres.Slides[0];

                // Create a full scale image
                Bitmap bmp = sld.GetThumbnail(1f, 1f);

                // Save the image to disk in JPEG format
                bmp.Save(dataDir + "Thumbnail_out.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            }
        }
    }
}