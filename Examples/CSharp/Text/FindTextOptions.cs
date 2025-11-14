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
This example demonstrates how to use TextSearchOptions to allows including text 
contained in slide notes when performing text replacement method.
*/

namespace CSharp.Text
{
    class FindTextOptions
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "TextOptionsExample.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "TextOptionsExample-out.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Set text search options
                TextSearchOptions options = new TextSearchOptions()
                {
                    IncludeNotes = true,
                    CaseSensitive = true
                };
                // Replace test
                pres.ReplaceText("old", "new", options, null);

                // Save result
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
