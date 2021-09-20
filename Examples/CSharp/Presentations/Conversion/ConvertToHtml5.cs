using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Slides.Export;
using Aspose.Slides.Export.Xaml;

/*
This example demonstrates the saving presentation in HTML5 operation.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ConvertToHtml5
    {
        public static void Run()
        {
            // The path to the documents directory
            string dataDir = RunExamples.GetDataDir_Conversion();

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "Demo.html");

            using (Presentation pres = new Presentation(dataDir + "Demo.pptx"))
            {
                // Export a presentation containing slides transitions, animations, and shapes animations to HTML5
                Html5Options options = new Html5Options()
                {
                    AnimateShapes = true,
                    AnimateTransitions = true
                };

                // Save presentation
                pres.Save(outFilePath, SaveFormat.Html5, options);
            }
        }
    }
}