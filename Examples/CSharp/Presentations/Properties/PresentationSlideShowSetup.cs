using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The example shows how to set up the slide show paramentes of the presentation.
After running this example, you can open the slideshow settings of the saved presentation and see the settings set by this example there.
*/
namespace CSharp.Presentations.Properties
{
    class PresentationSlideShowSetup
    {
        public static void Run()
        {
            //Path for out presentation
            string outPptxPath = Path.Combine(RunExamples.OutPath, "PresentationSlideShowSetup.pptx");

            using (var pres = new Presentation())
            {
                // Gets SlideShow settins
                var slideShow = pres.SlideShowSettings;

                // Sets "Using Timing" parameter
                slideShow.UseTimings = false;

                // Sets Pen Color
                var penColor = (ColorFormat)slideShow.PenColor;
                penColor.Color = Color.Green;

                // Adds slides for 
                {
                    pres.Slides.AddClone(pres.Slides[0]);
                    pres.Slides.AddClone(pres.Slides[0]);
                    pres.Slides.AddClone(pres.Slides[0]);
                    pres.Slides.AddClone(pres.Slides[0]);
                }

                // Sets Show Slide parameter
                slideShow.Slides = new SlidesRange()
                {
                    Start = 2,
                    End = 5
                };

                // Save presentation
                pres.Save(outPptxPath, SaveFormat.Pptx);
            }
        }
    }
}
