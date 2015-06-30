using Aspose.Slides.Util;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string sourceFile = @"E:\Aspose\Aspose Vs OpenXML\Aspose.Slides Vs OpenXML Presentation v1.1\Sample Files\Move a Paragraph from One Presentation to Another 1.pptx";
            string targetFile = @"E:\Aspose\Aspose Vs OpenXML\Aspose.Slides Vs OpenXML Presentation v1.1\Sample Files\Move a Paragraph from One Presentation to Another 2.pptx";
            MoveParagraphToPresentation(sourceFile, targetFile);
        }
        // Moves a paragraph range in a TextBody shape in the source document
        // to another TextBody shape in the target document.
        public static void MoveParagraphToPresentation(string sourceFile, string targetFile)
        {
            string Text = "";

            //Instantiate Presentation class that represents PPTX//Instantiate Presentation class that represents PPTX
            Presentation sourcePres = new Presentation(sourceFile);

            //Access first shape in first slide
            IShape shp = sourcePres.Slides[0].Shapes[0];
            if (shp.Placeholder != null)
            {
                //Get text from placeholder
                Text = ((IAutoShape)shp).TextFrame.Text;
                ((IAutoShape)shp).TextFrame.Text = "";
            }

            Presentation destPres = new Presentation(targetFile);
            //Access first shape in first slide
            IShape destshp = sourcePres.Slides[0].Shapes[0];
            if (destshp.Placeholder != null)
            {
                //Get text from placeholder
                ((IAutoShape)destshp).TextFrame.Text += Text;
            }

            sourcePres.Save(sourceFile, Export.SaveFormat.Pptx);
            destPres.Save(targetFile, Export.SaveFormat.Pptx);
        }
    }
}
