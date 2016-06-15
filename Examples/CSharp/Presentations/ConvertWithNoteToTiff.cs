using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class ConvertWithNoteToTiff
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "ConvertWithNoteToTiff.pptx"))
            {

                //Saving the presentation to TIFF notes
                pres.Save(dataDir + "TestNotes.tiff", Aspose.Slides.Export.SaveFormat.TiffNotes);
            }
        }
    }
}