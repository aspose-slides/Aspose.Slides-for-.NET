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
The following code sample demonstrates how to insert a new AudioFrame,  add captions to AudioFrame, remove all captions from a AudioFrame 
and extract captions from a AudioFrame instance and save them to a file.
*/

namespace CSharp.Slides.Media
{
    class AudioCaptionsExample
    {
        public static void Run()
        {
            string mediaFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "audio.mp3");
            string trackFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "bunny.vtt");
            string outCaption = Path.Combine(RunExamples.OutPath, "AudioCaption_out.vtt");
            string outAddPath = Path.Combine(RunExamples.OutPath, "AudioCaptionAdd_out.pptx");
            string outRemovePath = Path.Combine(RunExamples.OutPath, "AudioCaptionRemove_out.pptx");

            // Add captions to a VideoFrame
            using (Presentation pres = new Presentation())
            {
                IAudio audio = pres.Audios.AddAudio(File.ReadAllBytes(mediaFile));
                var audioFrame = pres.Slides[0].Shapes.AddAudioFrameEmbedded(10, 10, 50, 50, audio);

                // Adds the new captions track from file
                audioFrame.CaptionTracks.Add("New track", trackFile);

                pres.Save(outAddPath, SaveFormat.Pptx);
            }

            // Extract captions from a VideoFrame
            using (Presentation pres = new Presentation(outAddPath))
            {
                IAudioFrame audioFrame = pres.Slides[0].Shapes[0] as IAudioFrame;
                if (audioFrame != null)
                {
                    foreach (var captionTrack in audioFrame.CaptionTracks)
                    {
                        // Extracts the captions binary data and saves theme to the file
                        System.IO.File.WriteAllBytes(outCaption, captionTrack.BinaryData);
                    }

                    // Removes all captions from the VideoFrame
                    audioFrame.CaptionTracks.Clear();

                    pres.Save(outRemovePath, SaveFormat.Pptx);
                }
            }
        }
    }
}
