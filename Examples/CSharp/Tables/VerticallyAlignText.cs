using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class VerticallyAlignText
    {
        public static void Run()
        {
            // ExStart:VerticallyAlignText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Tables();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            // Get the first slide 
            ISlide slide = presentation.Slides[0];

            // Define columns with widths and rows with heights
            double[] dblCols = { 120, 120, 120, 120 };
            double[] dblRows = { 100, 100, 100, 100 };

            // Add table shape to slide
            ITable tbl = slide.Shapes.AddTable(100, 50, dblCols, dblRows);
            tbl[1, 0].TextFrame.Text = "10";
            tbl[2, 0].TextFrame.Text = "20";
            tbl[3, 0].TextFrame.Text = "30";

            // Accessing the text frame
            ITextFrame txtFrame = tbl[0, 0].TextFrame;

            // Create the Paragraph object for text frame
            IParagraph paragraph = txtFrame.Paragraphs[0];

            // Create Portion object for paragraph
            IPortion portion = paragraph.Portions[0];
            portion.Text = "Text here";
            portion.PortionFormat.FillFormat.FillType = FillType.Solid;
            portion.PortionFormat.FillFormat.SolidFillColor.Color = Color.Black;

            // Aligning the text vertically
            ICell cell = tbl[0, 0];
            cell.TextAnchorType = TextAnchorType.Center;
            cell.TextVerticalType = TextVerticalType.Vertical270;

            // Save Presentation
            presentation.Save(dataDir +  "Vertical_Align_Text_out.pptx", SaveFormat.Pptx);
            // ExEnd:VerticallyAlignText
         }
    }
}

