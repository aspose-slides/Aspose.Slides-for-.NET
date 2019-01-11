using System.IO;

using Aspose.Slides;
using System;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class MutilevelBullets
    {
        public static void Run()
        {
            //ExStart:MutilevelBullets
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
                ITextFrame text = aShp.AddTextFrame("");
                
                //clearing default paragraph
                text.Paragraphs.Clear();

                //Adding first paragraph
                IParagraph para1 = new Paragraph();
                para1.Text = "Content";
                para1.ParagraphFormat.Bullet.Type = BulletType.Symbol;
                para1.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);
                para1.ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillType.Solid;
                para1.ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.Black;
                //Setting bullet level
                para1.ParagraphFormat.Depth = 0;

                //Adding second paragraph
                IParagraph para2 = new Paragraph();
                para2.Text = "Second Level";
                para2.ParagraphFormat.Bullet.Type = BulletType.Symbol;
                para2.ParagraphFormat.Bullet.Char = '-';
                para2.ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillType.Solid;
                para2.ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.Black;
                //Setting bullet level
                para2.ParagraphFormat.Depth = 1;

                //Adding third paragraph
                IParagraph para3 = new Paragraph();
                para3.Text = "Third Level";
                para3.ParagraphFormat.Bullet.Type = BulletType.Symbol;
                para3.ParagraphFormat.Bullet.Char = Convert.ToChar(8226);
                para3.ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillType.Solid;
                para3.ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.Black;
                //Setting bullet level
                para3.ParagraphFormat.Depth = 2;

                //Adding fourth paragraph
                IParagraph para4 = new Paragraph();
                para4.Text = "Fourth Level";
                para4.ParagraphFormat.Bullet.Type = BulletType.Symbol;
                para4.ParagraphFormat.Bullet.Char = '-';
                para4.ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillType.Solid;
                para4.ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.Black;
                //Setting bullet level
                para4.ParagraphFormat.Depth = 3;

                //Adding paragraphs to collection
                text.Paragraphs.Add(para1);
                text.Paragraphs.Add(para2);
                text.Paragraphs.Add(para3);
                text.Paragraphs.Add(para4);

                //Writing the presentation as a PPTX file
                pres.Save(dataDir + "MultilevelBullet.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            
            }
            //ExEnd:MutilevelBullets
        }
    }
}