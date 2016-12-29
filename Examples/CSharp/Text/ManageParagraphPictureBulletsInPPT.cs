using System;
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

namespace Aspose.Slides.Examples.CSharp.Text
{
    class ManageParagraphPictureBulletsInPPT
    {
        public static void Run()
        {
            // ExStart:ManageParagraphPictureBulletsInPPT
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            Presentation presentation = new Presentation();

            // Accessing the first slide
            ISlide slide = presentation.Slides[0];

            // Instantiate the image for bullets
            Image image = new Bitmap(dataDir + "bullets.png");
            IPPImage ippxImage = presentation.Images.AddImage(image);

            // Adding and accessing Autoshape
            IAutoShape autoShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 400, 200);

            // Accessing the text frame of created autoshape
            ITextFrame textFrame = autoShape.TextFrame;

            // Removing the default exisiting paragraph
            textFrame.Paragraphs.RemoveAt(0);

            // Creating new paragraph
            Paragraph paragraph = new Paragraph();
            paragraph.Text = "Welcome to Aspose.Slides";

            // Setting paragraph bullet style and image
            paragraph.ParagraphFormat.Bullet.Type = BulletType.Picture;
            paragraph.ParagraphFormat.Bullet.Picture.Image = ippxImage;

            // Setting Bullet Height
            paragraph.ParagraphFormat.Bullet.Height = 100;

            // Adding Paragraph to text frame
            textFrame.Paragraphs.Add(paragraph);

            // Writing the presentation as a PPTX file
            presentation.Save(dataDir + "ParagraphPictureBulletsPPTX_out.pptx", SaveFormat.Pptx);
            // Writing the presentation as a PPT file
            presentation.Save(dataDir + "ParagraphPictureBulletsPPT_out.ppt", SaveFormat.Ppt);
            // ExEnd:ManageParagraphPictureBulletsInPPT
        }
    }
}