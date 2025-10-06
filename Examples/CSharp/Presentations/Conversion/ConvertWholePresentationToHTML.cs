using Aspose.Slides.Export;


/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Conversion
{
    public class ConvertWholePresentationToHTML
    {
        public static void Run()
        {
            //ExStart:ConvertWholePresentationToHTML
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "Convert_HTML.pptx"))
            {
                HtmlOptions htmlOpt = new HtmlOptions();
                htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

                NotesCommentsLayoutingOptions notesOptions = new NotesCommentsLayoutingOptions();
                notesOptions.NotesPosition = NotesPositions.BottomFull;
                htmlOpt.SlidesLayoutOptions = notesOptions;

                // Saving the presentation to HTML
                presentation.Save(dataDir + "ConvertWholePresentationToHTML_out.html", SaveFormat.Html, htmlOpt);
            }
            //ExEnd:ConvertWholePresentationToHTML
        }
    }
}