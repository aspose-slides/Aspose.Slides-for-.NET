using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class CloneWithinSamePresentationToEnd
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "CloneWithinSamePresentationToEnd.pptx"))
            {

                //Clone the desired slide to the end of the collection of slides in the same presentation
                ISlideCollection slds = pres.Slides;

                slds.AddClone(pres.Slides[0]);


                //Write the modified presentation to disk
                pres.Save(dataDir + "Aspose_cloned.pptx", SaveFormat.Pptx);

            }
        }
    }
}