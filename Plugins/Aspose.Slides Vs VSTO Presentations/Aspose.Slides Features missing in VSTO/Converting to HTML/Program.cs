using Aspose.Slides.Export;
using Aspose.Slides;

namespace Converting_to_HTML
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a presentation file
            Presentation pres = new Presentation(MyDir + "Conversion.ppt");

            HtmlOptions htmlOpt = new HtmlOptions();
            htmlOpt.HtmlFormatter = HtmlFormatter.CreateDocumentFormatter("", false);

            //Saving the presentation to HTML
            pres.Save(MyDir + "Converted.html", Aspose.Slides.Export.SaveFormat.Html, htmlOpt);
        }
    }
}
