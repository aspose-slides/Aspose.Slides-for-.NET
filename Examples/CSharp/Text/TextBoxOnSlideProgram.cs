using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class TextBoxOnSlideProgram
    {
        public static void Run()
        {
            // ExStart:TextBoxOnSlideProgram
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);
            
            // Instantiate PresentationEx// Instantiate PresentationEx
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add an AutoShape of Rectangle type
                IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50);

                // Add TextFrame to the Rectangle
                ashp.AddTextFrame(" ");

                // Accessing the text frame
                ITextFrame txtFrame = ashp.TextFrame;

                // Create the Paragraph object for text frame
                IParagraph para = txtFrame.Paragraphs[0];

                // Create Portion object for paragraph
                IPortion portion = para.Portions[0];

                // Set Text
                portion.Text = "Aspose TextBox";

                // Save the presentation to disk
                pres.Save(dataDir + "TextBox_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:TextBoxOnSlideProgram
        }
    }
}