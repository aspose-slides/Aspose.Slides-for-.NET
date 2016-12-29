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
    public class UpdatePresentationProperties
    {
        public static void Run()
        {
            //ExStart:UpdatePresentationProperties

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();

            // read the info of presentation 
            IPresentationInfo info = PresentationFactory.Instance.GetPresentationInfo(dataDir + "ModifyBuiltinProperties1.pptx");

            // obtain the current properties 
            IDocumentProperties props = info.ReadDocumentProperties();

            // set the new values of Author and Title fields 
            props.Author = "New Author";
            props.Title = "New Title";

            // update the presentation with a new values 
            info.UpdateDocumentProperties(props);
            info.WriteBindedPresentation(dataDir + "ModifyBuiltinProperties1.pptx");
            //ExEnd:UpdatePresentationProperties
        }
    }
}