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
This example shows how to specify the ZIP64 format to save a Presentation document.
*/

namespace CSharp.Presentations.Saving
{
    class SaveInZip64Mode
    {
        public static void Run()
        {
            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "PresentationZip64.pptx");

            // Create a new presentation
            using (Presentation pres = new Presentation())
            {
                // Save the presentation
                pres.Save(outFilePath, SaveFormat.Pptx, new PptxOptions()
                {
                    Zip64Mode = Zip64Mode.Always
                });
            }
        }
    }
}
