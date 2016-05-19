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
            string FilePath = @"..\..\..\Sample Files\";
            string FileName = FilePath + "Working With Size and Layout.pptx";
            
            //Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(FileName);
            Presentation auxPresentation = new Presentation();

            ISlide slide = presentation.Slides[0];

            //Set the slide size of generated presentations to that of source
            auxPresentation.SlideSize.Type = presentation.SlideSize.Type;
            auxPresentation.SlideSize.Size = presentation.SlideSize.Size;

            auxPresentation.Slides.InsertClone(0, slide);
            auxPresentation.Slides.RemoveAt(0);

            //Save Presentation to disk
            auxPresentation.Save(FileName, Aspose.Slides.Export.SaveFormat.Pptx);

        }
    }
}
