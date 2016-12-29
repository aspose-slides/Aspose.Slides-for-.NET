using System.IO;
using System.Text;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class ExportingHTMLText
    {
        public static void Run()
        {
            // ExStart:ExportingHTMLText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Load the presentation file
            using (Presentation pres = new Presentation(dataDir + "ExportingHTMLText.pptx"))
            {

                // Acesss the default first slide of presentation
                ISlide slide = pres.Slides[0];

                // Desired index
                int index = 0;

                // Accessing the added shape
                IAutoShape ashape = (IAutoShape)slide.Shapes[index];

                // Extracting first paragraph as HTML
                StreamWriter sw = new StreamWriter(dataDir + "output_out.html", false, Encoding.UTF8);

                //Writing Paragraphs data to HTML by providing paragraph starting index, total paragraphs to be copied
                sw.Write(ashape.TextFrame.Paragraphs.ExportToHtml(0, ashape.TextFrame.Paragraphs.Count, null));

                sw.Close();
            }
            // ExEnd:ExportingHTMLText
        }
    }
}