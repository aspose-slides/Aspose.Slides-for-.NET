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
    class HighlightText
    {
        public static void Run()
        {

            //ExStart:HighlightText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();
            Presentation presentation = new Presentation(dataDir +"SomePresentation.pptx");
            ((AutoShape)presentation.Slides[0].Shapes[0]).TextFrame.HighlightText("title", Color.LightBlue); // highlighting all words 'important'
            ((AutoShape)presentation.Slides[0].Shapes[0]).TextFrame.HighlightText("to", Color.Violet, new TextSearchOptions()
            {
                WholeWordsOnly = true
            }, null); // highlighting all separate 'the' occurrences
            presentation.Save(dataDir+ "SomePresentation-out2.pptx", SaveFormat.Pptx);

            //ExEnd:HighlightText
        }
    }
}
