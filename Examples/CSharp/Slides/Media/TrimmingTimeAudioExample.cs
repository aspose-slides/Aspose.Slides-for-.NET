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
The exampledemonstrates how to add a new audio frame with embedded audio and set the trimming times.
*/

namespace CSharp.Slides.Media
{
    class TrimmingTimeAudioExample
    {
        public static void Run()
        {
            string mediaFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "audio.m4a");
            string outPath = Path.Combine(RunExamples.OutPath, "AudioFrameTrim_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Add Audio Frame
                IAudio audio = pres.Audios.AddAudio(File.ReadAllBytes(mediaFile));
                IAudioFrame audioFrame = pres.Slides[0].Shapes.AddAudioFrameEmbedded(50, 50, 100, 100, audio);

                // Set the start trimming time 0.5 seconds
                audioFrame.TrimFromStart = 500f;

                // Set the end trimming time 1 seconds
                audioFrame.TrimFromEnd = 1000f;

                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
