using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class Convert_XPS
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "Convert_XPS.pptx"))
            {
                // Saving the presentation to XPS document
                pres.Save(dataDir + "XPS_Output_Without_XPSOption_out.xps", SaveFormat.Xps);
            }
        }
    }
}