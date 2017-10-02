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
    //ExStart:CustomHeaderAndFontsController
    public class CustomHeaderAndFontsController : EmbedAllFontsHtmlController
    {
        // Custom header template
        const string Header = +"<!DOCTYPE html>\n" +
                                "<html>\n" +
                                "<head>\n" +
                                "<meta http-equiv=\"Content-Type\" content=\"text/html; charset=UTF-8\">\n" +
                                "<meta http-equiv=\"X-UA-Compatible\" content=\"IE=9\">\n" +
                                "<link rel=\"stylesheet\" type=\"text/css\" href=\"{0}\">\n" +
                                "</head>";


        private readonly string m_cssFileName;

        public CustomHeaderAndFontsController(string cssFileName)
        {
            m_cssFileName = cssFileName;
        }

        public override void WriteDocumentStart(IHtmlGenerator generator, IPresentation presentation)
        {
            generator.AddHtml(string.Format(Header, m_cssFileName));
            WriteAllFonts(generator, presentation);
        }

        public override void WriteAllFonts(IHtmlGenerator generator, IPresentation presentation)
        {
            generator.AddHtml("<!-- Embedded fonts -->");
            base.WriteAllFonts(generator, presentation);
        }
    }
    //ExEnd:CustomHeaderAndFontsController
  
        
}
        
   
