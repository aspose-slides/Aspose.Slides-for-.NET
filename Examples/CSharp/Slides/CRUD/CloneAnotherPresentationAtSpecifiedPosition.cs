using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.CRUD
{
    class CloneAnotherPresentationAtSpecifiedPosition
    {
        public static void Run()
        {
            //ExStart:CloneAnotherPresentationAtSpecifiedPosition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // ExStart:CloneAnotherPresentationAtSpecifiedPosition
            // Instantiate Presentation class to load the source presentation file
            using (Presentation sourcePresentation = new Presentation(dataDir + "AccessSlides.pptx"))
            {
                // Instantiate Presentation class for destination presentation (where slide is to be cloned)
                using (Presentation destPres = new Presentation())
                {
                    // Clone the desired slide from the source presentation to the end of the collection of slides in destination presentation
                    ISlideCollection slideCollection = destPres.Slides;

                    // Clone the desired slide from the source presentation to the specified position in destination presentation
                    slideCollection.InsertClone(1, sourcePresentation.Slides[1]);

                    // ExEnd:CloneAnotherPresentationAtSpecifiedPosition
                    // Write the destination presentation to disk
                    destPres.Save(dataDir + "CloneAnotherPresentationAtSpecifiedPosition_out.pptx", SaveFormat.Pptx);
                }
            }
            //ExEnd:CloneAnotherPresentationAtSpecifiedPosition
        }
    }
}