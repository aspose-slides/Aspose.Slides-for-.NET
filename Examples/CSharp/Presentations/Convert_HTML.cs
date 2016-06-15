using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class Convert_HTML
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "Convert_HTML.pptx"))
            {
                HtmlOptions htmlOpt = new HtmlOptions();
                htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

                //Saving the presentation to HTML
                presentation.Save(dataDir + "demo.html", Aspose.Slides.Export.SaveFormat.Html, htmlOpt);
            }
        }
    }
}