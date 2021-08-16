using System.Drawing;
using System.IO;
using Aspose.Slides.Export;

/*
This example demonstrates how to export a Presentation to a Gif files.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    public class ConvertToGif
    {
        public static void Run()
        {
            // The path to the documents directory
            string dataDir = RunExamples.GetDataDir_Conversion();

            // The path to output file
            string outPath = Path.Combine(RunExamples.OutPath, "ConvertToGif.gif");

            // Instantiate a Presentation object that represents a presentation file
            Presentation presentation = new Presentation(dataDir + "ConvertToGif.pptx");

            // Save the presentation to Gif
            presentation.Save(outPath, SaveFormat.Gif, new GifOptions
            {
                FrameSize = new Size(540, 480), // the size of the resulted GIF  
                DefaultDelay = 1500, // how long each slide will be showed until it will be changed to the next one
                TransitionFps = 60 // increase FPS to better transition animation quality
            });
        }
    }
}