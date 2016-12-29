using System.Drawing;
using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class SetAutofitOftextframe
    {
        public static void Run()
        {
            // ExStart:SetAutofitOftextframe
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // ExStart:SetAutofitOftextframe
            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            // Access the first slide 
            ISlide slide = presentation.Slides[0];

            // Add an AutoShape of Rectangle type
            IAutoShape ashp = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 350, 350);

            // Add TextFrame to the Rectangle
            ashp.AddTextFrame(" ");
            ashp.FillFormat.FillType = FillType.NoFill;

            // Accessing the text frame
            ITextFrame txtFrame = ashp.TextFrame;
            txtFrame.TextFrameFormat.AutofitType = TextAutofitType.Shape;

            // Create the Paragraph object for text frame
            IParagraph para = txtFrame.Paragraphs[0];

            // Create Portion object for paragraph
            IPortion portion = para.Portions[0];
            portion.Text = "A quick brown fox jumps over the lazy dog. A quick brown fox jumps over the lazy dog.";
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.Black;

            // ExEnd:SetAutofitOftextframe
            // Save Presentation
            presentation.Save(dataDir + "formatText_out.pptx", SaveFormat.Pptx); 
        }
        // ExEnd:SetAutofitOftextframe
    }
}