using System;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.DOM.Ole;
using Aspose.Slides.Export;

/*
This example shows how to set or change video volume as a percentage. 
 */

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class VideoVolumeValueExample
    {
        public static void Run()
        {
            // Path to source presentation
            string videoFileName = Path.Combine(RunExamples.GetDataDir_Shapes(), "Wildlife.mp4");

            using (Presentation pres = new Presentation())
            {
                ISlide slide = pres.Slides[0];
                IVideo video = pres.Videos.AddVideo(File.ReadAllBytes(videoFileName));
                var videoFrame = slide.Shapes.AddVideoFrame(0, 0, 200, 200, video);

                // Set volume value
                videoFrame.VolumeValue = 75.5f;

                // Save presentation
                pres.Save(RunExamples.OutPath + "VideoVolumeValueExample-out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}