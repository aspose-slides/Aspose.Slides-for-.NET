using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.ActiveX
{
    public class LinkingVideoActiveXControl
    {
        public static void Run()
        {
            //ExStart:LinkingVideoActiveXControl
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_ActiveX();
            string dataVideo = RunExamples.GetDataDir_Video();

            // Instantiate Presentation class that represents PPTX file
            Presentation presentation = new Presentation(dataDir + "template.pptx");

            // Create empty presentation instance
            Presentation newPresentation = new Presentation();

            // Remove default slide
            newPresentation.Slides.RemoveAt(0);

            // Clone slide with Media Player ActiveX Control
            newPresentation.Slides.InsertClone(0, presentation.Slides[0]);

            // Access the Media Player ActiveX control and set the video path
            newPresentation.Slides[0].Controls[0].Properties["URL"] = dataVideo + "Wildlife.mp4";

            // Save the Presentation
            newPresentation.Save(dataDir + "LinkingVideoActiveXControl_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            //ExEnd:LinkingVideoActiveXControl
        }
    }
}