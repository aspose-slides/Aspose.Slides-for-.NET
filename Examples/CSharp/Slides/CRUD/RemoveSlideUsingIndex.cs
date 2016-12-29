using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.CRUD
{
    public class RemoveSlideUsingIndex
    {
        public static void Run()
        {
            //ExStart:RemoveSlideUsingIndex
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "RemoveSlideUsingIndex.pptx"))
            {

                // Removing a slide using its slide index
                pres.Slides.RemoveAt(0);

                // Writing the presentation file
                pres.Save(dataDir + "modified_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
            //ExEnd:RemoveSlideUsingIndex
        }
    }
}