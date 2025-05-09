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
This example shows how to export a presentation in the Handout layout to HTML5 document.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ConvertToHtml5Handout
    {
        public static void Run()
        {
            // The path to the documents directory
            string dataDir = RunExamples.GetDataDir_Conversion();

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "HandoutExample.html");

            using (Presentation pres = new Presentation(dataDir + "HandoutExample.pptx"))
            {
                // Set convertion options
                Html5Options options = new Html5Options
                {
                    SlidesLayoutOptions = new HandoutLayoutingOptions
                    {
                        Handout = HandoutType.Handouts4Horizontal
                    }
                };

                // Save result
                pres.Save(outFilePath, SaveFormat.Html5, options);
            }
        }
    }
}
