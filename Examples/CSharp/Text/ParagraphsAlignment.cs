using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class ParagraphsAlignment
    {
        public static void Run()
        {
            // ExStart:ParagraphsAlignment
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Instantiate a Presentation object that represents a PPTX file
            using (Presentation pres = new Presentation(dataDir + "ParagraphsAlignment.pptx"))
            {

                // Accessing first slide
                ISlide slide = pres.Slides[0];

                // Accessing the first and second placeholder in the slide and typecasting it as AutoShape
                ITextFrame tf1 = ((IAutoShape)slide.Shapes[0]).TextFrame;
                ITextFrame tf2 = ((IAutoShape)slide.Shapes[1]).TextFrame;

                // Change the text in both placeholders
                tf1.Text = "Center Align by Aspose";
                tf2.Text = "Center Align by Aspose";

                // Getting the first paragraph of the placeholders
                IParagraph para1 = tf1.Paragraphs[0];
                IParagraph para2 = tf2.Paragraphs[0];

                // Aligning the text paragraph to center
                para1.ParagraphFormat.Alignment = TextAlignment.Center;
                para2.ParagraphFormat.Alignment = TextAlignment.Center;

                //Writing the presentation as a PPTX file
                pres.Save(dataDir + "Centeralign_out.pptx", SaveFormat.Pptx);
            }
            // ExEnd:ParagraphsAlignment
        }
    }
}