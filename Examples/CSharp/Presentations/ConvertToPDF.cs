using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class ConvertToPDF
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            Presentation presentation = new Presentation(dataDir + "ConvertToPDF.pptx");
            
            //Save the presentation to PDF with default options
            presentation.Save(dataDir + "output.pdf", SaveFormat.Pdf);
                        
        }
    }
}