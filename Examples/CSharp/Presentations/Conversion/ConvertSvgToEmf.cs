using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
The following example demonstrates how to save the SVG image into a metafile.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertSvgToEmf
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            //Out path
            string resultPath = Path.Combine(RunExamples.OutPath, "SvgAsEmf.emf");

            // Creates the new SVG image
            ISvgImage svgImage = new SvgImage(System.IO.File.ReadAllText(dataDir + "content.svg"));

            // Saves the SVG image as a metafille
            using (var fileStream = System.IO.File.Create(resultPath))
            {
                svgImage.WriteAsEmf(fileStream);
            }
        }
    }
}
