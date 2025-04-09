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
The following code sample demonstrates how to insert a new VideoFrame, extract captions from a VideoFrame instance and save them to a file 
and add captions to it, and remove all captions from a VideoFrame.
*/

namespace CSharp.Slides.Media
{
    class VideoCaptionsExample
    {
        public static void Run()
        {
            string mediaFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "sample_bunny.mp4");
            string trackFile = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_Media(), "bunny.vtt");
            string outCaption = Path.Combine(RunExamples.OutPath, "Caption_out.vtt");
            string outAddPath = Path.Combine(RunExamples.OutPath, "VideoCaptionAdd_out.pptx");
            string outRemovePath = Path.Combine(RunExamples.OutPath, "VideoCaptionRemove_out.pptx");

            // Add captions to a VideoFrame
            using (Presentation pres = new Presentation())
            {
                IVideo video = pres.Videos.AddVideo(File.ReadAllBytes(mediaFile));
                var videoFrame = pres.Slides[0].Shapes.AddVideoFrame(0, 0, 100, 100, video);

                // Adds the new captions track from file
                videoFrame.CaptionTracks.Add("New track", trackFile);

                pres.Save(outAddPath, SaveFormat.Pptx);
            }

            // Extract captions from a VideoFrame
            using (Presentation pres = new Presentation(outAddPath))
            {
                IVideoFrame videoFrame = pres.Slides[0].Shapes[0] as VideoFrame;
                if (videoFrame != null)
                { 
                    foreach (var captionTrack in videoFrame.CaptionTracks)
                    {
                        // Extracts the captions binary data and saves theme to the file
                        System.IO.File.WriteAllBytes(outCaption, captionTrack.BinaryData);
                    }

                    // Removes all captions from the VideoFrame
                    videoFrame.CaptionTracks.Clear();

                    pres.Save(outRemovePath, SaveFormat.Pptx);
                }
            }
        }
    }
}
