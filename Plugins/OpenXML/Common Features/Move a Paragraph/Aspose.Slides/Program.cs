using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Move a Paragraph from One Presentation to Another 1.pptx";
            string DestFileName = FilePath + "Move a Paragraph from One Presentation to Another 2.pptx";
            MoveParagraphToPresentation(FileName, DestFileName);
        }
        // Moves a paragraph range in a TextBody shape in the source document
        // to another TextBody shape in the target document.
        public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
        {
            string Text = "";

            //Instantiate Presentation class that represents PPTX//Instantiate Presentation class that represents PPTX
            Presentation sourcePres = new Presentation(sourceFile);

            //Access first shape in first slide
            IShape shp = sourcePres.Slides[0].Shapes[0];
            if (shp.Placeholder != null)
            {
                //Get text from placeholder
                Text = ((IAutoShape)shp).TextFrame.Text;
                ((IAutoShape)shp).TextFrame.Text = "";
            }

            Presentation destPres = new Presentation(targetFile);
            //Access first shape in first slide
            IShape destshp = sourcePres.Slides[0].Shapes[0];
            if (destshp.Placeholder != null)
            {
                //Get text from placeholder
                ((IAutoShape)destshp).TextFrame.Text += Text;
            }

            sourcePres.Save(sourceFile, Aspose.Slides.Export.SaveFormat.Pptx);
            destPres.Save(targetFile, Aspose.Slides.Export.SaveFormat.Pptx);
        }
    }
}
