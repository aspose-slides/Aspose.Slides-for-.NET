using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class Convert_Tiff_Default
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "Convert_Tiff_Default.pptx"))
            {

                // Saving the presentation to TIFF document
                pres.Save(dataDir + "Tiff_out.tiff", Aspose.Slides.Export.SaveFormat.Tiff);
            }
        }
    }
}