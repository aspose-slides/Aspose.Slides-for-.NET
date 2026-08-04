using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following example shows how to render each paragraph in all AutoShapes on a slide as an image with custom scaling.
*/

namespace CSharp.Text
{
    class RenderParagraphExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "Presentation1.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                ISlide slide = pres.Slides[0];

                int shapeIndex = 0;
                foreach (IShape shape in slide.Shapes)
                {
                    shapeIndex++;

                    IAutoShape autoShape = shape as IAutoShape;
                    if (autoShape == null || autoShape.TextFrame == null)
                        continue;

                    int paragraphIndex = 0;
                    foreach (IParagraph paragraph in autoShape.TextFrame.Paragraphs)
                    {
                        paragraphIndex++;
                        string outFileName = Path.Combine(RunExamples.OutPath,
                            string.Format("shape{0}_paragraph{1}.png", shapeIndex, paragraphIndex));

                        using (IImage paragraphImage = paragraph.GetImage(2f, 2f))
                        {
                            paragraphImage?.Save(outFileName);
                        }
                    }
                }
            }
        }
    }
}