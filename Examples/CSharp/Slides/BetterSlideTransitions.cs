using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class BetterSlideTransitions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "BetterSlideTransitions.pptx"))
            {

                //Apply circle type transition on slide 1
                pres.Slides[0].SlideShowTransition.Type = TransitionType.Circle;


                //Set the transition time of 3 seconds
                pres.Slides[0].SlideShowTransition.AdvanceOnClick = true;
                pres.Slides[0].SlideShowTransition.AdvanceAfterTime = 3000;

                //Apply comb type transition on slide 2
                pres.Slides[1].SlideShowTransition.Type = TransitionType.Comb;


                //Set the transition time of 5 seconds
                pres.Slides[1].SlideShowTransition.AdvanceOnClick = true;
                pres.Slides[1].SlideShowTransition.AdvanceAfterTime = 5000;

                //Apply zoom type transition on slide 3
                pres.Slides[2].SlideShowTransition.Type = TransitionType.Zoom;


                //Set the transition time of 7 seconds
                pres.Slides[2].SlideShowTransition.AdvanceOnClick = true;
                pres.Slides[2].SlideShowTransition.AdvanceAfterTime = 7000;

                //Write the presentation to disk
                pres.Save(dataDir + "SampleTransition.pptx", SaveFormat.Pptx);

            }
        }
    }
}