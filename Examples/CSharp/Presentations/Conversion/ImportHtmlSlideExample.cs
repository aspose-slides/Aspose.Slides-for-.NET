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
The following code sample demonstrates how to insert HTML content into the presentation slide collection, starting from the empty space on the slide with index equal to 0.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ImportHtmlSlideExample
    {
        public static void Run()
        {
            // The path to the documents directory
            string dataDir = RunExamples.GetDataDir_Conversion();

            // The path to html document
            string htmlFileName = Path.Combine(dataDir, "TestHtml.html");

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "OutputConvertedHtml.pptx");

            using (var presentation = new Presentation())
            {
                using (Stream inputStream = new FileStream(htmlFileName, FileMode.Open))
                    presentation.Slides.InsertFromHtml(0, inputStream, true);

                // Save the presentation
                presentation.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
