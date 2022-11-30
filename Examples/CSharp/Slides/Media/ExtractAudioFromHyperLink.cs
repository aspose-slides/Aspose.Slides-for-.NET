using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;

/*
Next example demonstrates how to get the embedded audio file from HyperlinkClick settings using the "Sound" property.
 */

namespace CSharp.Slides.Media
{
    class ExtractAudioFromHyperLink
    {
        public static void Run()
        {
            string pptxFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "HyperlinkSound.pptx");
            string outMediaPath = Path.Combine(RunExamples.OutPath, "HyperlinkSound.mpg");

            using (Presentation pres = new Presentation(pptxFile))
            {
                // Gets the first shape hyperlink
                IHyperlink link = pres.Slides[0].Shapes[0].HyperlinkClick;

                if (link.Sound != null)
                {
                    // Extracts the hyperlink sound in byte array
                    byte[] audioData = link.Sound.BinaryData;

                    // Saves effect sound to media file
                    File.WriteAllBytes(outMediaPath, audioData);
                }
            }
        }
    }
}
