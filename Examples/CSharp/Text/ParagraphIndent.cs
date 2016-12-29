using System.IO;

using Aspose.Slides;
using System;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class ParagraphIndent
    {
        public static void Run()
        {
            // ExStart:ParagraphIndent

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Presentation Class
            Presentation pres = new Presentation();

            // Get first slide
            ISlide sld = pres.Slides[0];

            // Add a Rectangle Shape
            IAutoShape rect = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 500, 150);

            // Add TextFrame to the Rectangle
            ITextFrame tf = rect.AddTextFrame("This is first line \rThis is second line \rThis is third line");

            // Set the text to fit the shape
            tf.TextFrameFormat.AutofitType = TextAutofitType.Shape;

            // Hide the lines of the Rectangle
            rect.LineFormat.FillFormat.FillType = FillType.Solid;

            // Get first Paragraph in the TextFrame and set its Indent
            IParagraph para1 = tf.Paragraphs[0];
            // Setting paragraph bullet style and symbol
            para1.ParagraphFormat.Bullet.Type = BulletType.Symbol;
            para1.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);
            para1.ParagraphFormat.Alignment = TextAlignment.Left;

            para1.ParagraphFormat.Depth = 2;
            para1.ParagraphFormat.Indent = 30;

            // Get second Paragraph in the TextFrame and set its Indent
            IParagraph para2 = tf.Paragraphs[1];
            para2.ParagraphFormat.Bullet.Type = BulletType.Symbol;
            para2.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);
            para2.ParagraphFormat.Alignment = TextAlignment.Left;
            para2.ParagraphFormat.Depth = 2;
            para2.ParagraphFormat.Indent = 40;

            // Get third Paragraph in the TextFrame and set its Indent
            IParagraph para3 = tf.Paragraphs[2];
            para3.ParagraphFormat.Bullet.Type = BulletType.Symbol;
            para3.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);
            para3.ParagraphFormat.Alignment = TextAlignment.Left;
            para3.ParagraphFormat.Depth = 2;
            para3.ParagraphFormat.Indent = 50;

            //Write the Presentation to disk
            pres.Save(dataDir + "InOutDent_out.pptx", SaveFormat.Pptx);
            // ExEnd:ParagraphIndent            
        }
    }
}