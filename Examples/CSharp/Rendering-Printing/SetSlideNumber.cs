using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Rendering.Printing
{
    class SetSlideNumber
    {
        public static void Run()
        {
            //ExStart:SetSlideNumber
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Rendering();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "HelloWorld.pptx"))
            {
                // Get the slide number
                int firstSlideNumber = presentation.FirstSlideNumber;

                // Set the slide number
                presentation.FirstSlideNumber=10;

                presentation.Save(dataDir + "Set_Slide_Number_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SetSlideNumber
        }
    }
}

