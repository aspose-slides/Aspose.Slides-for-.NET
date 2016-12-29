using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class AccessBuiltinProperties
    {
        public static void Run()
        {
            //ExStart:AccessBuiltinProperties

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();

            // Instantiate the Presentation class that represents the presentation
            Presentation pres = new Presentation(dataDir + "AccessBuiltin Properties.pptx");

            // Create a reference to IDocumentProperties object associated with Presentation
            IDocumentProperties documentProperties = pres.DocumentProperties;

            // Display the builtin properties
            System.Console.WriteLine("Category : " + documentProperties.Category);
            System.Console.WriteLine("Current Status : " + documentProperties.ContentStatus);
            System.Console.WriteLine("Creation Date : " + documentProperties.CreatedTime);
            System.Console.WriteLine("Author : " + documentProperties.Author);
            System.Console.WriteLine("Description : " + documentProperties.Comments);
            System.Console.WriteLine("KeyWords : " + documentProperties.Keywords);
            System.Console.WriteLine("Last Modified By : " + documentProperties.LastSavedBy);
            System.Console.WriteLine("Supervisor : " + documentProperties.Manager);
            System.Console.WriteLine("Modified Date : " + documentProperties.LastSavedTime);
            System.Console.WriteLine("Presentation Format : " + documentProperties.PresentationFormat);
            System.Console.WriteLine("Last Print Date : " + documentProperties.LastPrinted);
            System.Console.WriteLine("Is Shared between producers : " + documentProperties.SharedDoc);
            System.Console.WriteLine("Subject : " + documentProperties.Subject);
            System.Console.WriteLine("Title : " + documentProperties.Title);
            //ExEnd:AccessBuiltinProperties            
        }
    }
}