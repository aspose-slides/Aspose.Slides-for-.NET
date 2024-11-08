using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
The following example shows how to set the grid spacing to 72 points (1 inch) and save the PowerPoint presentation.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class GridProperties
    {
        public static void Run()
        {
            string outFilePath = Path.Combine(RunExamples.OutPath, "GridProperties-out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Set grid spacing
                pres.ViewProperties.GridSpacing = 72f;

                // Save presentation
                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
