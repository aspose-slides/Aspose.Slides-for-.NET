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
    public class ThumbnailFromSlideInNotes
    {
        public static void Run()
        {
            //ExStart:ThumbnailFromSlideInNotes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Thumbnail();

            // Instantiate a Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "ThumbnailFromSlideInNotes.pptx"))
            {
                // Access the first slide
                ISlide sld = pres.Slides[0];

                // User defined dimension
                int desiredX = 1200;
                int desiredY = 800;

                // Getting scaled value  of X and Y
                float ScaleX = (float)(1.0 / pres.SlideSize.Size.Width) * desiredX;
                float ScaleY = (float)(1.0 / pres.SlideSize.Size.Height) * desiredY;

               
                // Create a full scale image                
                IImage img = sld.GetImage(ScaleX, ScaleY);
                // Save the image to disk in JPEG format
                img.Save(dataDir + "Notes_tnail_out.jpg", ImageFormat.Jpeg);
            }
            //ExEnd:ThumbnailFromSlideInNotes
        }
    }
}