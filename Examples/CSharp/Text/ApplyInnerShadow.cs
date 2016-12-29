using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class ApplyInnerShadow
    {
        public static void Run()
        {
            // ExStart:ApplyInnerShadow
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
                pres.Save(dataDir + "ApplyInnerShadow_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:ApplyInnerShadow
        }
    }
}
