using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides.Examples.CSharp.Conversion
{
    //ExStart:LinkAllFontsHtmlController
    class LinkAllFontsHtmlController : EmbedAllFontsHtmlController
    {
        private readonly string m_basePath;

        public LinkAllFontsHtmlController(string[] fontNameExcludeList, string basePath)
            : base(fontNameExcludeList)
        {
            m_basePath = basePath;
        }

        public override void WriteFont(
            IHtmlGenerator generator,
            IFontData originalFont,
            IFontData substitutedFont,
            string fontStyle,
            string fontWeight,
            byte[] fontData)
        {
            string fontName = substitutedFont == null ? originalFont.FontName : substitutedFont.FontName;
            string path = string.Format("{0}.woff", fontName); // some path sanitaze may be needed
            File.WriteAllBytes(Path.Combine(m_basePath, path), fontData);

            generator.AddHtml("<style>");
            generator.AddHtml("@font-face { ");
            generator.AddHtml(string.Format("font-family: '{0}'; ", fontName));
            generator.AddHtml(string.Format("src: url('{0}')", path));

            generator.AddHtml(" }");
            generator.AddHtml("</style>");
        }
    
    }
  //  ExEnd:LinkAllFontsHtmlController
}
