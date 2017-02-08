using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Layout
{
    class SetSizeAndType
    {
        public static void Run()
        {
            //ExStart:SetSizeAndType
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Layout();

            // ExStart:SettSizeAndType
            // Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(dataDir + "AccessSlides.pptx");
            Presentation auxPresentation = new Presentation();

            ISlide slide = presentation.Slides[0];

            // Set the slide size of generated presentations to that of source
            auxPresentation.SlideSize.Type = presentation.SlideSize.Type;
            auxPresentation.SlideSize.Size = presentation.SlideSize.Size;

            auxPresentation.Slides.InsertClone(0, slide);
            auxPresentation.Slides.RemoveAt(0);
            // Save Presentation to disk
            auxPresentation.Save(dataDir + "Set_Size&Type_out.pptx", SaveFormat.Pptx);
            //ExEnd:SetSizeAndType
        }
    }
}