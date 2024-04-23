using System.Drawing;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Background
{
    public class SetImageAsBackground
    {
        public static void Run()
        {
            //ExStart:SetImageAsBackground
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Background();

            // Instantiate the Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "SetImageAsBackground.pptx"))
            {

                // Set the background with Image
                pres.Slides[0].Background.Type = BackgroundType.OwnBackground;
                pres.Slides[0].Background.FillFormat.FillType = FillType.Picture;
                pres.Slides[0].Background.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;

                // Set the picture
                IImage img = Images.FromFile(dataDir + "Tulips.jpg");

                // Add image to presentation's images collection
                IPPImage imgx = pres.Images.AddImage(img);

                pres.Slides[0].Background.FillFormat.PictureFillFormat.Picture.Image = imgx;

                // Write the presentation to disk
                pres.Save(dataDir + "ContentBG_Img_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SetImageAsBackground
        }
    }
}