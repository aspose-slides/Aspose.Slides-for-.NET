using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Animation;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/*
The example shows how to extracting an audio file from a slide timeline.
*/
namespace CSharp.Slides.Media
{
    class ExtractAudioFromTimeline
    {
        public static void Run()
        {
            string pptxFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "AnimationAudio.pptx");
            string outMediaPath = Path.Combine(RunExamples.OutPath, "MediaTimeline.mpg");

            using (Presentation pres = new Presentation(pptxFile))
            {
                // Gets first slide of the presentation
                ISlide slide = pres.Slides[0];

                // Gets the effects sequence for the slide
                ISequence effectsSequence = slide.Timeline.MainSequence;

                // Extracts the effect sound in byte array
                byte[] audio = effectsSequence[0].Sound.BinaryData;

                // Saves effect sound to media file
                File.WriteAllBytes(outMediaPath, audio);
            }
        }
    }
}
