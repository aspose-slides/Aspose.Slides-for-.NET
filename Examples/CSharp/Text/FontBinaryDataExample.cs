using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
The following example shows how to retrieve binary font data from a presentation.
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class FontBinaryDataExample
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation(dataDir + "Presentation.pptx"))
            {
                // Retrieve all fonts used in the presentation
                IFontData[] fonts = pres.FontsManager.GetFonts();

                // Get the byte array representing the regular style of the first font in the presentation
                byte[] bytes = pres.FontsManager.GetFontBytes(fonts[0], FontStyle.Regular);

                // The path to output file
                string outFilePath = Path.Combine(RunExamples.OutPath, fonts[0].FontName + ".ttf");

                // Save font
                File.WriteAllBytes(outFilePath, bytes);
            }
        }
    }
}
