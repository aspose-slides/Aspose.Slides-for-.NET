using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to use PreferRegisteredFontsOverEmbedded property to specifies 
whether registered fonts (system or loaded) should be preferred over embedded fonts 
with the same font name and style.
*/

namespace CSharp.Text
{
    class PreferRegisteredFontsOverEmbedded
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Text(), "EmbeddedFonts.pptx");

            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "presentation-out.pdf");

            using (Presentation pres = new Presentation(presentationName))
            {
                pres.FontsManager.PreferRegisteredFontsOverEmbedded = true;
                pres.Save(resultPath, SaveFormat.Pdf);
            }
        }
    }
}