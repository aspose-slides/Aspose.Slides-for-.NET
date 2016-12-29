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
    public class ChangePosition
    {
        public static void Run()
        {
            //ExStart:ChangePosition
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_CRUD();

            // Instantiate Presentation class to load the source presentation file
            using (Presentation pres = new Presentation(dataDir + "ChangePosition.pptx"))
            {
                // Get the slide whose position is to be changed
                ISlide sld = pres.Slides[0];

                // Set the new position for the slide
                sld.SlideNumber = 2;

                // Write the presentation to disk
                pres.Save(dataDir + "Aspose_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:ChangePosition
        }
    }
}