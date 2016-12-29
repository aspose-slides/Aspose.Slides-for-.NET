using System.Drawing;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Thumbnail
{
    public class ThumbnailFromSlide
    {
        public static void Run()
        {
            // ExStart:ThumbnailFromSlide
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Thumbnail();

            // Instantiate a Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "ThumbnailFromSlide.pptx"))
            {

                // Access the first slide
                ISlide sld = pres.Slides[0];

                // Create a full scale image
                Bitmap bmp = sld.GetThumbnail(1f, 1f);

                // Save the image to disk in JPEG format
                bmp.Save(dataDir + "Thumbnail_out.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            }
            // ExEnd:ThumbnailFromSlide
        }
    }
}