using System;
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
    class AddCustomDocumentProperties
    {
        public static void Run()
        {
            //ExStart:AddCustomDocumentProperties
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();

            // Instantiate the Presentation class
            Presentation presentation = new Presentation();

            // Getting Document Properties
            IDocumentProperties documentProperties = presentation.DocumentProperties;

            // Adding Custom properties
            documentProperties["New Custom"] = 12;
            documentProperties["My Name"] = "Mudassir";
            documentProperties["Custom"] = 124;

            // Getting property name at particular index
            String getPropertyName = documentProperties.GetCustomPropertyName(2);

            // Removing selected property
            documentProperties.RemoveCustomProperty(getPropertyName);

            // Saving presentation
            presentation.Save(dataDir + "CustomDocumentProperties_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            //ExEnd:AddCustomDocumentProperties
        }
    }
}
