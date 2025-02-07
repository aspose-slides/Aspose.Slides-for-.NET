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
The example demonstrates how to add a new audio frame with embedded audio and set its volume to 85%.
*/

namespace CSharp.Slides.Media
{
    class VolumeAudioExample
    {
        public static void Run()
        {
            string mediaFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "audio.m4a");
            string outPath = Path.Combine(RunExamples.OutPath, "AudioFrameValue_out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Add Audio Frame
                IAudio audio = pres.Audios.AddAudio(File.ReadAllBytes(mediaFile));
                IAudioFrame audioFrame = pres.Slides[0].Shapes.AddAudioFrameEmbedded(50, 50, 100, 100, audio);

                // Set the audio volume for 85%
                audioFrame.VolumeValue = 85f;

                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}
