using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example shows how to set the transparency property for tables in a presentation.
*/

namespace Aspose.Slides.Examples.CSharp.Tables
{
    class TableTransparency
    {
        public static void Run()
        {
            // The path to the presentation
            string presentationFilePat = RunExamples.GetDataDir_Tables() + "TableTransparency.pptx";

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "TableTransparency_out.pptx");

            // Create a new presentation
            using (Presentation pres = new Presentation(presentationFilePat))
            {
                // Get a table
                ITable table = (ITable)pres.Slides[0].Shapes[1];

                //Set transparency of the table to 62%
                table.TableFormat.Transparency = 0.62f; 

                // Save presentation
                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
