using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class GetTextFrameFormatEffectiveData
    {
        public static void Run() {

            //ExStart:GetTextFrameFormatEffectiveData

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            using (Presentation pres = new Presentation(dataDir + "Presentation1.pptx"))
            {
                IAutoShape shape = pres.Slides[0].Shapes[0] as IAutoShape;

                ITextFrameFormat textFrameFormat = shape.TextFrame.TextFrameFormat;
                ITextFrameFormatEffectiveData effectiveTextFrameFormat = textFrameFormat.GetEffective();


                Console.WriteLine("Anchoring type: " + effectiveTextFrameFormat.AnchoringType);
                Console.WriteLine("Autofit type: " + effectiveTextFrameFormat.AutofitType);
                Console.WriteLine("Text vertical type: " + effectiveTextFrameFormat.TextVerticalType);
                Console.WriteLine("Margins");
                Console.WriteLine("   Left: " + effectiveTextFrameFormat.MarginLeft);
                Console.WriteLine("   Top: " + effectiveTextFrameFormat.MarginTop);
                Console.WriteLine("   Right: " + effectiveTextFrameFormat.MarginRight);
                Console.WriteLine("   Bottom: " + effectiveTextFrameFormat.MarginBottom);

            }
            //ExEnd:GetTextFrameFormatEffectiveData

        }
    }
}
