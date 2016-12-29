using System.Drawing;
using Aspose.Slides.Export;
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
    class AddRelativeScaleHeightPictureFrame
    {
        public static void Run()
        {
            //ExStart:AddRelativeScaleHeightPictureFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Instantiate presentation object
            using (Presentation presentation = new Presentation())
            {

                // Load Image to be added in presentaiton image collection
                Image img = new Bitmap(dataDir + "aspose-logo.jpg");
                IPPImage image = presentation.Images.AddImage(img);

                // Add picture frame to slide
                IPictureFrame pf = presentation.Slides[0].Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 50, 100, 100, image);

                // Setting relative scale width and height
                pf.RelativeScaleHeight = 0.8f;
                pf.RelativeScaleWidth = 1.35f;

                // Save presentation
                presentation.Save(dataDir + "Adding Picture Frame with Relative Scale_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddRelativeScaleHeightPictureFrame
        }
    }
}

