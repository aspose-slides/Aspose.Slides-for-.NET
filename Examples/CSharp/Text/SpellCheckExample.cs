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
This example demonstrates how to use SpellCheck property to enable or disable spell checking 
for individual text portions within a presentation.
*/

namespace CSharp.Text
{
    class SpellCheckExample
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "SpellChecksExample.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "SpellChecksExample-out.pptx");

            using (var pres = new Presentation(presentationName))
            {
                // Access the first portion of text inside the first shape on the first slide
                var portion = ((AutoShape)pres.Slides[0].Shapes[0]).TextFrame.Paragraphs[0].Portions[0];

                // Read spell checking property
                Console.WriteLine("SpellCheck is {0}", portion.PortionFormat.SpellCheck);

                // Disable spell checking for this text portion
                portion.PortionFormat.SpellCheck = false;

                // Save the modified presentation
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}