using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class GetTextStyleEffectiveData
    {
        public static void Run() {

            //ExStart:
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();


            using (Presentation pres = new Presentation(dataDir + "Presentation1.pptx"))
            {
                IAutoShape shape = pres.Slides[0].Shapes[0] as IAutoShape;

                ITextStyleEffectiveData effectiveTextStyle = shape.TextFrame.TextFrameFormat.TextStyle.GetEffective();

                for (int i = 0; i <= 8; i++)
                {
                    IParagraphFormatEffectiveData effectiveStyleLevel = effectiveTextStyle.GetLevel(i);
                    Console.WriteLine("= Effective paragraph formatting for style level #" + i + " =");

                    Console.WriteLine("Depth: " + effectiveStyleLevel.Depth);
                    Console.WriteLine("Indent: " + effectiveStyleLevel.Indent);
                    Console.WriteLine("Alignment: " + effectiveStyleLevel.Alignment);
                    Console.WriteLine("Font alignment: " + effectiveStyleLevel.FontAlignment);
                }

            }

            //ExEnd:
        }

    }
}
