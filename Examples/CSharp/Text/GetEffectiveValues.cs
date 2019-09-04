using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class GetEffectiveValues
    {
        public static void Run() {

            //ExStart:GetEffectiveValues
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation(dataDir + "Presentation1.pptx"))
            {
                IAutoShape shape = pres.Slides[0].Shapes[0] as IAutoShape;

                ITextFrameFormat localTextFrameFormat = shape.TextFrame.TextFrameFormat;
                ITextFrameFormatEffectiveData effectiveTextFrameFormat = localTextFrameFormat.GetEffective();

                IPortionFormat localPortionFormat = shape.TextFrame.Paragraphs[0].Portions[0].PortionFormat;
                IPortionFormatEffectiveData effectivePortionFormat = localPortionFormat.GetEffective();
            }

            //ExEnd:GetEffectiveValues


        }
    }
}
