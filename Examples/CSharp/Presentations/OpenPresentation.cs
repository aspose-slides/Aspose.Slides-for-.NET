 
using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class OpenPresentation
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Opening the presentation file by passing the file path to the constructor of Presentation class
            Presentation pres = new Presentation(dataDir + "OpenPresentation.pptx");

            // Printing the total number of slides present in the presentation
            System.Console.WriteLine(pres.Slides.Count.ToString());
        }
    }
}