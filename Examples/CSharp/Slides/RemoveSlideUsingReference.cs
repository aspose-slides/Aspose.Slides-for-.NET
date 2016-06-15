using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class RemoveSlideUsingReference
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "RemoveSlideUsingReference.pptx"))
            {

                //Accessing a slide using its index in the slides collection
                ISlide slide = pres.Slides[0];


                //Removing a slide using its reference
                pres.Slides.Remove(slide);


                //Writing the presentation file
                pres.Save(dataDir + "modified.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}