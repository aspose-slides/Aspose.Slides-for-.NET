using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
The following example shows how to get the embedding level for a font.
*/

namespace Aspose.Slides.Examples.CSharp.Text
{
    class FontEmbeddingLevelExample
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            using (Presentation pres = new Presentation(dataDir + "Presentation.pptx"))
            {
                // Retrieve all fonts used in the presentation
                IFontData[] fontDatas = pres.FontsManager.GetFonts();

                // Get the byte array representing the regular style of the first font in the presentation
                byte[] bytes = pres.FontsManager.GetFontBytes(fontDatas[0], FontStyle.Regular);

                // Determine the embedding level of the font
                EmbeddingLevel embeddingLevel = pres.FontsManager.GetFontEmbeddingLevel(bytes, fontDatas[0].FontName);

                // Print embedding level to console
                Console.WriteLine("Font \"{0}\" has \"{1}\" Embedding Level", fontDatas[0].FontName, embeddingLevel);
            }
        }
    }
}
