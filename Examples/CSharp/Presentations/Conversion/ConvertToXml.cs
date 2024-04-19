using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example demonstrates export a presentation in PowerPoint XML format 
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertToXml
    {
        public static void Run()
        {
            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "Example.xml");

            // Create a new presentation
            using (Presentation pres = new Presentation())
            {
                // Save presentation in XML
                pres.Save(outFilePath, SaveFormat.Xml);
            }
        }
    }
}
