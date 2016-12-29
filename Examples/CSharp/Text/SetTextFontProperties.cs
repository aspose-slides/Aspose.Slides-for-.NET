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

namespace Aspose.Slides.Examples.CSharp.Text
{
    class SetTextFontProperties
    {
        public static void Run()
        {
            // ExStart:SetTextFontProperties
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Instantiate Presentation
            using (Presentation presentation = new Presentation())
            {
                // ExStart:SetTextFontProperties

                // Get first slide
                ISlide sld = presentation.Slides[0];

                // Add an AutoShape of Rectangle type
                IAutoShape ashp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 50, 200, 50);

                // Remove any fill style associated with the AutoShape
                ashp.FillFormat.FillType = FillType.NoFill;

                // Access the TextFrame associated with the AutoShape
                ITextFrame tf = ashp.TextFrame;
                tf.Text = "Aspose TextBox";

                // Access the Portion associated with the TextFrame
                IPortion port = tf.Paragraphs[0].Portions[0];

                // Set the Font for the Portion
                port.PortionFormat.LatinFont = new FontData("Times New Roman");

                // Set Bold property of the Font
                port.PortionFormat.FontBold = NullableBool.True;

                // Set Italic property of the Font
                port.PortionFormat.FontItalic = NullableBool.True;

                // Set Underline property of the Font
                port.PortionFormat.FontUnderline = TextUnderlineType.Single;

                // Set the Height of the Font
                port.PortionFormat.FontHeight = 25;

                // Set the color of the Font
                port.PortionFormat.FillFormat.FillType = FillType.Solid;
                port.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue;

                // ExEnd:SetTextFontProperties
                // Write the PPTX to disk 
                presentation.Save(dataDir + "SetTextFontProperties_out.pptx", SaveFormat.Pptx);
            }
            // ExEnd:SetTextFontProperties
        }
    }
}
