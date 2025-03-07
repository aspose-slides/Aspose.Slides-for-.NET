using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following code samples demonstrate how to save a PPTX presentation without hyperlinks containing JavaScript calls.
*/

namespace CSharp.Presentations.Saving
{
    class JavaScriptLinkExample
    {
        public static void Run()
        {
            //Path for source presentation
            string pptxFile = Path.Combine(RunExamples.GetDataDir_PresentationSaving(), "JavaScriptLink.pptx");
            //Out path
            string resultPath = Path.Combine(RunExamples.OutPath, "JavaScriptLink-out.pptx");

            using (Presentation pres = new Presentation(pptxFile))
            {
                pres.Save(resultPath, SaveFormat.Pptx, new PptxOptions()
                {
                    SkipJavaScriptLinks = true
                });
            }
        }
    }
}
