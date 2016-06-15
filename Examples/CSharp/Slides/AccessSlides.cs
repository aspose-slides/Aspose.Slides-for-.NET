using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class AccessSlides
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "AccessSlides.pptx"))
            {

                //Accessing a slide using its slide index
                ISlide slide = pres.Slides[0];

                System.Console.WriteLine("Slide Number: " + slide.SlideNumber);

            }
        }
    }
}