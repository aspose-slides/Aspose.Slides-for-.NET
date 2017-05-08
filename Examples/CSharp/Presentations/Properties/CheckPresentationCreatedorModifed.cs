using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/


namespace Aspose.Slides.Examples.CSharp.Presentations.Properties
{
    public class CheckPresentationCreatedorModifed
    {
        public static void Run()
        {
            //ExStart:CheckPresentationCreatedorModifed

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();

            IPresentationInfo info =
                  PresentationFactory.Instance.GetPresentationInfo(Path.Combine(RootFolder, "props.pptx"));

            IDocumentProperties props = info.ReadDocumentProperties();

            string app = props.NameOfApplication;
            string ver = props.AppVersion;
            //ExEnd:CheckPresentationCreatedorModifed
        }
    }
}