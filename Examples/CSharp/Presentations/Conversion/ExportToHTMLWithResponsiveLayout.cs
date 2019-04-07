using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Conversion
{
    class ExportToHTMLWithResponsiveLayout
    {
        public static void Run() {

            //ExStart:ExportToHTMLWithResponsiveLayout
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            Presentation presentation = new Presentation(dataDir+"SomePresentation.pptx");
            HtmlOptions saveOptions = new HtmlOptions();
            saveOptions.SvgResponsiveLayout = true;
            presentation.Save(dataDir+"SomePresentation-out.html", SaveFormat.Html, saveOptions);
            //ExEnd:ExportToHTMLWithResponsiveLayout
        }
    }
}
