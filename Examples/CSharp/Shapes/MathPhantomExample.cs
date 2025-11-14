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
The example shows how to use the MathPhantom that represent a phantom math object (<m:phant>). 
Phantom math object  affects the layout of its child element without necessarily displaying it. 
A phantom can hide its base expression while preserving its width, height, or depth - useful for aligning formulas 
or reserving space. Visibility and geometry behavior are controlled by properties such as Show, ZeroWid, ZeroAsc, ZeroDesc, 
nd Transp.
*/

namespace CSharp.Shapes
{
    class MathPhantomExample
    {
        public static void Run()
        {
            //Path for output presentation
            string outPptxFile = Path.Combine(RunExamples.OutPath, "MathPhantom_out.pptx");

            using (Presentation pres = new Presentation())
            {
                IAutoShape autoShape = pres.Slides[0].Shapes.AddMathShape(0, 0, 500, 50);
                IMathParagraph mathParagraph =
                    ((MathPortion)autoShape.TextFrame.Paragraphs[0].Portions[0]).MathParagraph;

                var eq1 = new MathematicalText("eq1");
                var eq2 = new MathematicalText("eq2");
                // Create phantom math object
                var phant = new MathPhantom(new MathFraction(new MathematicalText("1"), new MathematicalText("2")))
                    { Show = false, ZeroAsc = true };
                var first = new MathematicalText("    (1)");
                var sect = new MathematicalText("    (2)");
                var second = new MathematicalText().Join(phant).Join(sect);
                var nums = new MathArray(new IMathElement[] { first, second });
                var eqs = new MathDelimiter(new MathArray(new IMathElement[] { eq1, eq2 }))
                    { BeginningCharacter = '{', EndingCharacter = '\0' };
                var wholeBlock = new MathematicalText().Join(eqs).Join(" ").Join(nums);
                mathParagraph.Add(wholeBlock);

                pres.Save(outPptxFile, SaveFormat.Pptx);
            }
        }
    }
}
