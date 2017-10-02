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
    public class ConvertingPresentationToHTMLWithPreservingOriginalFonts
    {
        public static void Run()
        {
            //ExStart:ConvertingPresentationToHTMLWithPreservingOriginalFonts
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            using (Presentation pres = new Presentation("input.pptx"))
            {
                // exclude default presentation fonts
                string[] fontNameExcludeList = { "Calibri", "Arial" };

                EmbedAllFontsHtmlController embedFontsController = new EmbedAllFontsHtmlController(fontNameExcludeList);

                HtmlOptions htmlOptionsEmbed = new HtmlOptions
                {
                    HtmlFormatter = HtmlFormatter.CreateCustomFormatter(embedFontsController)
                };

                pres.Save("input-PFDinDisplayPro-Regular-installed.html", SaveFormat.Html, htmlOptionsEmbed);
            }
            //ExEnd:ConvertingPresentationToHTMLWithPreservingOriginalFonts
        }
    }
}