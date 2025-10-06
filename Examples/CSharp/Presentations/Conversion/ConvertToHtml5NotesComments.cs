using System.IO;
using Aspose.Slides.Export;

// This example demonstrates export a presentation with notes to HTML5 format.
// Note: Please use the license to get the correct output.

namespace Aspose.Slides.Examples.CSharp.Conversion
{
    public class ConvertToHtml5NotesComments
    {
        public static void Run()
        {
            // Path to source directory
            string dataDir = RunExamples.GetDataDir_Conversion();
            string resultPath = Path.Combine(RunExamples.OutPath, "Html5NotesResult.html");

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "ConvertWithNote.pptx"))
            {
                // Save a result
                pres.Save(resultPath, SaveFormat.Html5, new Html5Options()
                {
                    OutputPath = RunExamples.OutPath,
                    SlidesLayoutOptions = new NotesCommentsLayoutingOptions() { NotesPosition = NotesPositions.BottomTruncated }
                });
            }
        }
    }
}