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
This example demonstrates a using of GetChildren() method of the IMathElement interface.
*/
namespace CSharp.Shapes
{
    class MathShape_GetChildren
    {
        public static void Run()
        {
            //Path for output presentation
            string outPptxFile = Path.Combine(RunExamples.OutPath, "MathShape_GetChildren_out.pptx");

            using (var presentation = new Presentation())
            {
                // Get first slide
                ISlide slide = presentation.Slides[0];

                // Create MathShape in the first slide
                IAutoShape mathShape = slide.Shapes.AddMathShape(10, 10, 500, 500);
                // Create MathParagraph
                IMathParagraph mathParagraph = (mathShape.TextFrame.Paragraphs[0].Portions[0] as MathPortion).MathParagraph;

                // Create MathBlock
                IMathBlock mathBlock = new MathBlock(new MathematicalText("F").Join("+").Join(new MathematicalText("1").Divide("y")).Underbar());

                // Add MathBlock to the MathParagraph
                mathParagraph.Add(mathBlock);
                
                // Print all elements of the mathBlock
                ForEachMathElement(mathBlock);

                presentation.Save(outPptxFile, SaveFormat.Pptx);
            }
        }

        private static void ForEachMathElement(IMathElement root)
        {
            foreach (IMathElement child in root.GetChildren())
            {
                Console.WriteLine(child.GetType() + (child is MathematicalText ? " : " +((MathematicalText)child).Value : ""));

                //recursive
                ForEachMathElement(child);
            }
        }
    }
}
