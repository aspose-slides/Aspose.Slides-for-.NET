using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class RemoveSlideUsingIndex
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "RemoveSlideUsingIndex.pptx"))
            {

                // Removing a slide using its slide index
                pres.Slides.RemoveAt(0);


                //Writing the presentation file
                pres.Save(dataDir + "modified_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
        }
    }
}