using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Opening
{
    public class OpenPresentation
    {
        public static void Run()
        {
            //ExStart:OpenPresentation
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationOpening();

            // Opening the presentation file by passing the file path to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "OpenPresentation.pptx");

            // Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString());
            //ExStart:OpenPresentation
        }
    }
}