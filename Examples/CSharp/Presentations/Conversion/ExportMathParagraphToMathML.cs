using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.MathText;

// This example demonstrates export a mathematical paragraph or block to MathML format. 

namespace CSharp.Presentations.Conversion
{
    class ExportMathParagraphToMathML
    {
        public static void Run()
        {
            string outSvgFileName = Path.Combine(RunExamples.OutPath, "mathml.xml");

            using (Presentation pres = new Presentation())
            {
                var autoShape = pres.Slides[0].Shapes.AddMathShape(0, 0, 500, 50);
                var mathParagraph = ((MathPortion) autoShape.TextFrame.Paragraphs[0].Portions[0]).MathParagraph;

                mathParagraph.Add(new MathematicalText("a").SetSuperscript("2").Join("+")
                    .Join(new MathematicalText("b").SetSuperscript("2")).Join("=")
                    .Join(new MathematicalText("c").SetSuperscript("2")));

                using (Stream stream = new FileStream(outSvgFileName, FileMode.Create))
                    mathParagraph.WriteAsMathMl(stream);
            }
        }
    }
}