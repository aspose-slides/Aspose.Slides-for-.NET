using System.Drawing;
using System.Drawing.Imaging;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class CreateShapeThumbnail
    {
        public static void Run()
        {
            //ExStart:CreateShapeThumbnail
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Instantiate a Presentation class that represents the presentation file
            using (Presentation presentation = new Presentation(dataDir + "HelloWorld.pptx"))
            {
                // Create a full scale image
                using (Bitmap bitmap = presentation.Slides[0].Shapes[0].GetThumbnail())
                {
                    // Save the image to disk in PNG format
                    bitmap.Save(dataDir + "Shape_thumbnail_out.png", ImageFormat.Png);
                }
            }
            //ExEnd:CreateShapeThumbnail
        }
    }
}


