using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class TextBoxHyperlink
    {
        public static void Run()
        {
            // ExStart:TextBoxHyperlink
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate a Presentation class that represents a PPTX
            Presentation pptxPresentation = new Presentation();

            // Get first slide
            ISlide slide = pptxPresentation.Slides[0];

            // Add an AutoShape of Rectangle Type
            IShape pptxShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 150, 150, 50);

            // Cast the shape to AutoShape
            IAutoShape pptxAutoShape = (IAutoShape)pptxShape;

            // Access ITextFrame associated with the AutoShape
            pptxAutoShape.AddTextFrame("");

            ITextFrame ITextFrame = pptxAutoShape.TextFrame;

            // Add some text to the frame
            ITextFrame.Paragraphs[0].Portions[0].Text = "Aspose.Slides";

            // Set Hyperlink for the portion text
            IHyperlinkManager HypMan = ITextFrame.Paragraphs[0].Portions[0].PortionFormat.HyperlinkManager;
            HypMan.SetExternalHyperlinkClick("http://www.aspose.com");
            // Save the PPTX Presentation
            pptxPresentation.Save(dataDir + "hLinkPPTX_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            // ExEnd:TextBoxHyperlink
        }
    }
}