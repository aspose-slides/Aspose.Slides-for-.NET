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
    public class CloneAtEndOfAnotherSpecificPosition
    {
        public static void Run()
        {
            //ExStart:CloneAtEndOfAnotherSpecificPosition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // Instantiate Presentation class to load the source presentation file
            using (Presentation srcPres = new Presentation(dataDir + "CloneAtEndOfAnotherSpecificPosition.pptx"))
            {
                // Instantiate Presentation class for destination presentation (where slide is to be cloned)
                using (Presentation destPres = new Presentation())
                {
                    // Clone the desired slide from the source presentation to the end of the collection of slides in destination presentation
                    ISlideCollection slds = destPres.Slides;

                    // Clone the desired slide from the source presentation to the specified position in destination presentation
                    slds.InsertClone(1, srcPres.Slides[1]);

                    // Write the destination presentation to disk
                    destPres.Save(dataDir + "Aspose1_out.pptx", SaveFormat.Pptx);
                }
            }
            //ExEnd:CloneAtEndOfAnotherSpecificPosition
        }
    }
}