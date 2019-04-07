using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Hyperlinks
{
    class SetHyperLinkColor
    {
        public static void Run() {

            //ExStart:SetHyperLinkColor
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Hyperlink();
            using (Presentation presentation = new Presentation())
            {
                IAutoShape shape1 = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 450, 50, false);
                shape1.AddTextFrame("This is a sample of colored hyperlink.");
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.HyperlinkClick = new Hyperlink("https://www.aspose.com/");
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.HyperlinkClick.ColorSource = HyperlinkColorSource.PortionFormat;
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.FillFormat.FillType = FillType.Solid;
                shape1.TextFrame.Paragraphs[0].Portions[0].PortionFormat.FillFormat.SolidFillColor.Color = Color.Red;

                IAutoShape shape2 = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 200, 450, 50, false);
                shape2.AddTextFrame("This is a sample of usual hyperlink.");
                shape2.TextFrame.Paragraphs[0].Portions[0].PortionFormat.HyperlinkClick = new Hyperlink("https://www.aspose.com/");

                presentation.Save(dataDir+"presentation-out-hyperlink.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SetHyperLinkColor
        }
    }
}
