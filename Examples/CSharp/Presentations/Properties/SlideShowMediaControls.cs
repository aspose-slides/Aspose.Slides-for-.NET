using System.IO;
using Aspose.Slides.Export;

/*
This example shows how to enable/disable slideshow media controls in a presentation.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class SlideShowMediaControls
    {
        public static void Run()
        {
            // Path to HTML document
            string outFilePath = Path.Combine(RunExamples.OutPath, "SlideShowMediaControl.pptx");

            using (Presentation pres = new Presentation())
            {
                // Еnable media control display in slideshow mode. 
                pres.SlideShowSettings.ShowMediaControls = true;
                
                // Save presentation in HTML5 format.
                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}

