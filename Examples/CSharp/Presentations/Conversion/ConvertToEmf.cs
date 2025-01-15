using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example demonstrates how to convert the first slide from a PowerPoint presentation into a metafile.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertToEmf
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            //Out path
            string resultPath = Path.Combine(RunExamples.OutPath, "Result.emf");

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation presentation = new Presentation(dataDir + "HelloWorld.pptx"))
            {
                using (Stream fileStream = System.IO.File.Create(resultPath))
                {
                    // Saves the first slide as a metafille
                    presentation.Slides[0].WriteAsEmf(fileStream);
                }
            }
        }
    }
}
