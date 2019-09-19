using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class SetCustomBulletsNumber
    {
        public static void Run() {

            //ExStart:SetCustomBulletsNumber

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (var presentation = new Presentation())
            {
                var shape = presentation.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 200, 200, 400, 200);

                // Accessing the text frame of created autoshape
                ITextFrame textFrame = shape.TextFrame;

                // Removing the default exisiting paragraph
                textFrame.Paragraphs.RemoveAt(0);

                // First list
                var paragraph1 = new Paragraph { Text = "bullet 2" };
                paragraph1.ParagraphFormat.Depth = 4; 
                paragraph1.ParagraphFormat.Bullet.NumberedBulletStartWith = 2;
                paragraph1.ParagraphFormat.Bullet.Type = BulletType.Numbered;
                textFrame.Paragraphs.Add(paragraph1);

                var paragraph2 = new Paragraph { Text = "bullet 3" };
                paragraph2.ParagraphFormat.Depth = 4;
                paragraph2.ParagraphFormat.Bullet.NumberedBulletStartWith = 3; 
                paragraph2.ParagraphFormat.Bullet.Type = BulletType.Numbered;  
                textFrame.Paragraphs.Add(paragraph2);

                
                var paragraph5 = new Paragraph { Text = "bullet 7" };
                paragraph5.ParagraphFormat.Depth = 4;
                paragraph5.ParagraphFormat.Bullet.NumberedBulletStartWith = 7;
                paragraph5.ParagraphFormat.Bullet.Type = BulletType.Numbered;
                textFrame.Paragraphs.Add(paragraph5);

                presentation.Save(dataDir + "SetCustomBulletsNumber-slides.pptx", SaveFormat.Pptx);
            }


            //ExEnd:SetCustomBulletsNumber

        }
    }
}
