using System;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System.IO;
using Aspose.Slides.MathText;

// This example demonstrates export a mathematical paragraph or block to Latex format. 

namespace CSharp.Presentations.Conversion
{
    class ExportMathParagraphToLatex
    {
        public static void Run()
        {
            // Create a new presentation
            using (Presentation pres = new Presentation())
            {
                // Add a math shape.
                var autoShape = pres.Slides[0].Shapes.AddMathShape(0, 0, 500, 50);

                // Get a math paragraph.
                var mathParagraph = ((MathPortion)autoShape.TextFrame.Paragraphs[0].Portions[0]).MathParagraph;

                // Add a formula.
                mathParagraph.Add(new MathematicalText("a").SetSuperscript("2").Join("+")
                    .Join(new MathematicalText("b").SetSuperscript("2")).Join("=")
                    .Join(new MathematicalText("c").SetSuperscript("2")));

                // Get formula string in Latex format.
                string latexString = mathParagraph.ToLatex();

                // Output the resulting Latex string to the console.
                Console.WriteLine("Latex representation of a math paragraph: \"" + latexString + "\"");
            }
        }
    }
}