
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class CloneAtEndOfAnother
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate Presentation class to load the source presentation file
            using (Presentation srcPres = new Presentation(dataDir + "CloneAtEndOfAnother.pptx"))
            {
                //Instantiate Presentation class for destination PPTX (where slide is to be cloned)
                using (Presentation destPres = new Presentation())
                {
                    //Clone the desired slide from the source presentation to the end of the collection of slides in destination presentation
                    ISlideCollection slds = destPres.Slides;

                    slds.AddClone(srcPres.Slides[0]);

                    //Write the destination presentation to disk
                    destPres.Save(dataDir + "Aspose_out.pptx", SaveFormat.Pptx);
                }
            }
        }
    }
}