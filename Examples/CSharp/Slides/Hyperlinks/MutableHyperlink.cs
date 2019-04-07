using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Hyperlinks
{
    class MutableHyperlink
    {

        public static void Run() {

            //ExStart:MutableHyperlink
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Hyperlink();

            using (Presentation presentation = new Presentation())
            {
                
                IAutoShape shape1 = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 600, 50, false);
                shape1.AddTextFrame("Aspose: File Format APIs");
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.HyperlinkClick = new Hyperlink("https://www.aspose.com/");
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.HyperlinkClick.Tooltip = "More than 70% Fortune 100 companies trust Aspose APIs";
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 32;

                presentation.Save(dataDir+"presentation-out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:MutableHyperlink

        }
    }
}
