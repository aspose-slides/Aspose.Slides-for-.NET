 
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class PPTtoPPTX
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a PPTX file
            Presentation pres = new Presentation(dataDir + "PPTtoPPTX.ppt");

            // Saving the PPTX presentation to PPTX format
            pres.Save(dataDir + "PPTtoPPTX_out.pptx", SaveFormat.Pptx);


        }
    }
}