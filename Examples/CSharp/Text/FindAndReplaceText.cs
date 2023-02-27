using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides.LowCode;

/*
This example demonstrates a search for all portions of “[this block] " in the presentation and then replaces them with “my text” filled in red.

NOTE. To obtain the correct result, a valid aspose.slides license must be used.
*/

namespace CSharp.Text
{
    class FindAndReplaceText
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "TextReplaceExample.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "TextReplaceExample-out.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                PortionFormat format = new PortionFormat
                {
                    FontHeight = 24f,
                    FontItalic = NullableBool.True,
                    FillFormat =
                    {
                        FillType = FillType.Solid,
                        SolidFillColor =
                        {
                            Color = Color.Red
                        }
                    }
                };
                Aspose.Slides.Util.SlideUtil.FindAndReplaceText(pres, true, "[this block] ", "my text", format);
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
