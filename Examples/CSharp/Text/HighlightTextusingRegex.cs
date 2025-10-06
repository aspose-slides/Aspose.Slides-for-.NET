using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

/*
The example demonstrates how to highlight text in TextFrame using Regex.
*/

namespace CSharp.Text
{
    class HighlightTextUsingRegx
    {
        public static void Run() {

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Path to output file
            string outFileName = Path.Combine(RunExamples.OutPath, "omePresentation-out.pptx");

            using (Presentation presentation = new Presentation(dataDir + "SomePresentation.pptx"))
            {
                ((AutoShape)presentation.Slides[0].Shapes[0]).TextFrame.HighlightRegex(new Regex(@"\b[^\s]{5,}\b"), Color.Blue, null);
                presentation.Save(outFileName, SaveFormat.Pptx);
            }
        }
    }
}
