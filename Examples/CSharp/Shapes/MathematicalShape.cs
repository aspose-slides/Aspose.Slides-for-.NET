using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.MathText;

/*
This example demonstrates of using API for creation a mathematical expression for Pythagorean theorem.
*/
namespace CSharp.Shapes
{
    class MathematicalShape
    {
        public static void Run()
        {
            //Path for output presentation
            string outpptxFile = Path.Combine(RunExamples.OutPath, "MathematicalShape_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create a new AutoShape of the type Rectangle to host mathematical content inside and adds it to the end of the collection.
                IAutoShape mathShape = pres.Slides[0].Shapes.AddMathShape(10, 10, 100, 25);

                // Cteate mathematical paragraph that is a container for mathematical blocks.
                IMathParagraph mathParagraph = ((MathPortion)mathShape.TextFrame.Paragraphs[0].Portions[0]).MathParagraph;

                // Create mathematical expression as an instance of mathematical text that contained within a MathParagraph.
                IMathBlock mathBlock = new MathematicalText("c")
                    .SetSuperscript("2")
                    .Join("=")
                    .Join(new MathematicalText("a")
                        .SetSuperscript("2"))
                    .Join("+")
                    .Join(new MathematicalText("b")
                        .SetSuperscript("2"));

                // Add mathematical expression to the mathematical paragraph.
                mathParagraph.Add(mathBlock);

                pres.Save(outpptxFile, SaveFormat.Pptx); 
            }
        }
    }
}
