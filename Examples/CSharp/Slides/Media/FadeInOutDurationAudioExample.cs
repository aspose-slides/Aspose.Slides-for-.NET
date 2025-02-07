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
using Aspose.Slides.Export;

/*
The example shows how to add a new audio frame with embedded audio and configure its fade-in and fade-out durations.
*/

namespace CSharp.Slides.Media
{
    class FadeInOutDurationAudioExample
    {
        public static void Run()
        {
            string mediaFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "audio.m4a");
            string outPath = Path.Combine(RunExamples.OutPath, "AudioFrameFade_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Add Audio Frame
                IAudio audio = pres.Audios.AddAudio(File.ReadAllBytes(mediaFile));
                IAudioFrame audioFrame = pres.Slides[0].Shapes.AddAudioFrameEmbedded(50, 50, 100, 100, audio);

                // Set the duration of the starting fade for 200ms
                audioFrame.FadeInDuration = 200f;
                // Set the duration of the ending fade for 500ms
                audioFrame.FadeOutDuration = 500f;

                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
