using Aspose.Slides;
using Aspose.Slides.Export;
using Aspose.Slides.SlideShow;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Managing_Slides_Transitions
{
    class Program
    {
        static void Main(string[] args)
        {
            string Path = @"E:\Aspose\Aspose Vs OpenXML\Files\";

            //Instantiate Presentation class that represents a presentation file
            using (Presentation pres = new Presentation(Path + "Sample.pptx"))
            {

                //Apply circle type transition on slide 1
                pres.Slides[0].SlideShowTransition.Type = TransitionType.Circle;

                //Apply comb type transition on slide 2
                pres.Slides[1].SlideShowTransition.Type = TransitionType.Comb;

                //Apply zoom type transition on slide 3
                pres.Slides[2].SlideShowTransition.Type = TransitionType.Zoom;

                //Write the presentation to disk
                pres.Save(Path + "SampleTransition.pptx", SaveFormat.Pptx);

            }
        }
    }
}
