using Aspose.Slides.Examples.CSharp.Conversion;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ConvertingPresentationToHtmlWithEmbedAllFontsHtmlController
    {
        public static void Run()
        {
            //ExStart:ConvertingPresentationToHtmlWithEmbedAllFontsHtmlController
            string dataDir = RunExamples.GetDataDir_Conversion();
            using (Presentation pres = new Presentation(dataDir+"presentation.pptx"))
            {
                // exclude default presentation fonts
                string[] fontNameExcludeList = { };


                Paragraph para = new Paragraph();

                EmbedAllFontsHtmlController embedFontsController = new EmbedAllFontsHtmlController(fontNameExcludeList);

                LinkAllFontsHtmlController linkcont = new LinkAllFontsHtmlController(fontNameExcludeList, @"C:\Windows\Fonts\");

                HtmlOptions htmlOptionsEmbed = new HtmlOptions
                {
                    HtmlFormatter = HtmlFormatter.CreateCustomFormatter(linkcont)
                };

                pres.Save("pres.html", SaveFormat.Html, htmlOptionsEmbed);
                //ExEnd:ConvertingPresentationToHtmlWithEmbedAllFontsHtmlController
            }
        }
    }
}