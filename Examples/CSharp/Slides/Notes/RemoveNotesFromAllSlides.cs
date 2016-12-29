using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Notes
{
    class RemoveNotesFromAllSlides
    {
        public static void Run()
        {
            //ExStart:RemoveNotesFromAllSlides
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Notes();

            // ExStart:RemoveNotesFromAllSlides
            // Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(dataDir + "AccessSlides.pptx");

            // Removing notes of all slides
            INotesSlideManager mgr = null;
            for (int i = 0; i < presentation.Slides.Count; i++)
            {
                mgr = presentation.Slides[i].NotesSlideManager;
                mgr.RemoveNotesSlide();
            }
            // Save presentation to disk
            presentation.Save(dataDir + "RemoveNotesFromAllSlides_out.pptx", SaveFormat.Pptx);
            // ExEnd:RemoveNotesFromAllSlides
        }
    }
}