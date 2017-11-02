using Aspose.Slides.Export;
using Aspose.Slides.Charts;
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
    class AddingSuperscriptAndSubscriptTextInTextFrame
    {
        public static void Run()
        {
           
             //ExStart:AddingSuperscriptAndSubscriptTextInTextFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            using (Presentation presentation = new Presentation(dataDir+"test.pptx"))
            {
                // Get slide
                ISlide slide = presentation.Slides[0];

                // Create text box
                IAutoShape shape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
                ITextFrame textFrame = shape.TextFrame;
                textFrame.Paragraphs.Clear();

                // Create paragraph for superscript text
                IParagraph superPar = new Paragraph();

                // Create portion with usual text
                IPortion portion1 = new Portion();
                portion1.Text = "SlideTitle";
                superPar.Portions.Add(portion1);

                // Create portion with superscript text
                IPortion superPortion = new Portion();
                superPortion.PortionFormat.Escapement = 30;
                superPortion.Text = "TM";
                superPar.Portions.Add(superPortion);

                // Create paragraph for subscript text
                IParagraph paragraph2 = new Paragraph();

                // Create portion with usual text
                IPortion portion2 = new Portion();
                portion2.Text = "a";
                paragraph2.Portions.Add(portion2);

                // Create portion with subscript text
                IPortion subPortion = new Portion();
                subPortion.PortionFormat.Escapement = -25;
                subPortion.Text = "i";
                paragraph2.Portions.Add(subPortion);

                // Add paragraphs to text box
                textFrame.Paragraphs.Add(superPar);
                textFrame.Paragraphs.Add(paragraph2);

                presentation.Save(dataDir+"TestOut.pptx", SaveFormat.Pptx);
                System.Diagnostics.Process.Start(dataDir + "TestOut.pptx");
             } 
            //ExEnd:AddingSuperscriptAndSubscriptTextInTextFrame
        }
    }
}