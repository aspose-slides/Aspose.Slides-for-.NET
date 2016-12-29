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
    public class UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate
    {
        //ExStart:UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();

            DocumentProperties template;
            IPresentationInfo info = PresentationFactory.Instance.GetPresentationInfo(dataDir + "template.pptx");
            template = (DocumentProperties)info.ReadDocumentProperties();

            template.Author = "Template Author";
            template.Title = "Template Title";
            template.Category = "Template Category";
            template.Keywords = "Keyword1, Keyword2, Keyword3";
            template.Company = "Our Company";
            template.Comments = "Created from template";
            template.ContentType = "Template Content";
            template.Subject = "Template Subject";

            UpdateByTemplate(dataDir + "doc1.pptx", template);
            UpdateByTemplate(dataDir + "doc2.odp", template);
            UpdateByTemplate(dataDir + "doc3.ppt", template);
        }

        private static void UpdateByTemplate(string path, IDocumentProperties template)
        {
            IPresentationInfo toUpdate = PresentationFactory.Instance.GetPresentationInfo(path);
            toUpdate.UpdateDocumentProperties(template);
            toUpdate.WriteBindedPresentation(path);
        }
        //ExStart:UpdatePresentationPropertiesUsingPropertiesOfAnotherPresentationAsATemplate
    }
}