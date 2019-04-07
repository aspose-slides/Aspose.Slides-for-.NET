using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Text
{
    class HighlightTextUsingRegx
    {
        public static void Run() {

            //ExStart:HighlightTextUsingRegx
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            Presentation presentation = new Presentation(dataDir + "SomePresentation.pptx");
            TextHighlightingOptions options = new TextHighlightingOptions();
            ((AutoShape)presentation.Slides[0].Shapes[0]).TextFrame.HighlightRegex(@"\b[^\s]{5,}\b", Color.Blue, options); // highlighting all words with 10 symbols or longer
            presentation.Save(dataDir+ "SomePresentation-out.pptx", SaveFormat.Pptx);

            //ExEnd:HighlightTextUsingRegx
        }
    }
}
