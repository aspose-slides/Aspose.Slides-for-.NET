using System.Drawing;
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
    class ManageParagraphFontProperties
    {
        public static void Run()
        {
            // ExStart:ManageParagraphFontProperties
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Instantiate PresentationEx 
            using (Presentation presentation = new Presentation(dataDir + "DefaultFonts.pptx"))
            {
                // Accessing a slide using its slide position
                ISlide slide = presentation.Slides[0];
                
                // Accessing the first and second placeholder in the slide and typecasting it as AutoShape
                ITextFrame tf1 = ((IAutoShape)slide.Shapes[0]).TextFrame;
                ITextFrame tf2 = ((IAutoShape)slide.Shapes[1]).TextFrame;
                
                // Accessing the first Paragraph
                IParagraph para1 = tf1.Paragraphs[0];
                IParagraph para2 = tf2.Paragraphs[0];

                // Justify the paragraph
                para2.ParagraphFormat.Alignment = TextAlignment.JustifyLow;

                // Accessing the first portion
                IPortion port1 = para1.Portions[0];
                IPortion port2 = para2.Portions[0];

                // Define new fonts
                FontData fd1 = new FontData("Elephant");
                FontData fd2 = new FontData("Castellar");

                // Assign new fonts to portion
                port1.PortionFormat.LatinFont = fd1;
                port2.PortionFormat.LatinFont = fd2;

                // Set font to Bold
                port1.PortionFormat.FontBold = NullableBool.True;
                port2.PortionFormat.FontBold = NullableBool.True;

                // Set font to Italic
                port1.PortionFormat.FontItalic = NullableBool.True;
                port2.PortionFormat.FontItalic = NullableBool.True;

                // Set font color
                port1.PortionFormat.FillFormat.FillType = FillType.Solid;
                port1.PortionFormat.FillFormat.SolidFillColor.Color = Color.Purple;
                port2.PortionFormat.FillFormat.FillType = FillType.Solid;
                port2.PortionFormat.FillFormat.SolidFillColor.Color = Color.Peru;

                // Write the PPTX to disk 
                presentation.Save(dataDir + "ManagParagraphFontProperties_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:ManageParagraphFontProperties
        }
    }
}
