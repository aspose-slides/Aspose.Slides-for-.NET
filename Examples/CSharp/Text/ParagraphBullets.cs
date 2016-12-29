using System.IO;

using Aspose.Slides;
using System;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class ParagraphBullets
    {
        public static void Run()
        {
            // ExStart:ParagraphBullets
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Creating a presenation instance
            using (Presentation pres = new Presentation())
            {

                // Accessing the first slide
                ISlide slide = pres.Slides[0];


                // Adding and accessing Autoshape
                IAutoShape aShp = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 400, 200);

                // Accessing the text frame of created autoshape
                ITextFrame txtFrm = aShp.TextFrame;

                // Removing the default exisiting paragraph
                txtFrm.Paragraphs.RemoveAt(0);

                // Creating a paragraph
                Paragraph para = new Paragraph();

                // Setting paragraph bullet style and symbol
                para.ParagraphFormat.Bullet.Type = BulletType.Symbol;
                para.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);

                // Setting paragraph text
                para.Text = "Welcome to Aspose.Slides";

                // Setting bullet indent
                para.ParagraphFormat.Indent = 25;

                // Setting bullet color
                para.ParagraphFormat.Bullet.Color.ColorType = ColorType.RGB;
                para.ParagraphFormat.Bullet.Color.Color = Color.Black;
                para.ParagraphFormat.Bullet.IsBulletHardColor = NullableBool.True; // set IsBulletHardColor to true to use own bullet color

                // Setting Bullet Height
                para.ParagraphFormat.Bullet.Height = 100;

                // Adding Paragraph to text frame
                txtFrm.Paragraphs.Add(para);

                // Creating second paragraph
                Paragraph para2 = new Paragraph();

                // Setting paragraph bullet type and style
                para2.ParagraphFormat.Bullet.Type = BulletType.Numbered;
                para2.ParagraphFormat.Bullet.NumberedBulletStyle = NumberedBulletStyle.BulletCircleNumWDBlackPlain;

                // Adding paragraph text
                para2.Text = "This is numbered bullet";

                // Setting bullet indent
                para2.ParagraphFormat.Indent = 25;

                para2.ParagraphFormat.Bullet.Color.ColorType = ColorType.RGB;
                para2.ParagraphFormat.Bullet.Color.Color = Color.Black;
                para2.ParagraphFormat.Bullet.IsBulletHardColor = NullableBool.True; // set IsBulletHardColor to true to use own bullet color

                // Setting Bullet Height
                para2.ParagraphFormat.Bullet.Height = 100;

                // Adding Paragraph to text frame
                txtFrm.Paragraphs.Add(para2);


                //Writing the presentation as a PPTX file
                pres.Save(dataDir + "Bullet_out.pptx", SaveFormat.Pptx);

            }
            // ExEnd:ParagraphBullets
        }
    }
}