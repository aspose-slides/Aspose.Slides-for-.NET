using System;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.DOM.Ole;
using Aspose.Slides.Export;

/*
This sample demonstrates how to set the trimming start and end time.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class SetVideoTrimFromEnd
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

                // sets the trimming start time to 12sec
                videoFrame.TrimFromStart = 12000f;

                // sets the triming end time to 16sec
                videoFrame.TrimFromEnd = 14000f;

                // Save presentation
                pres.Save(RunExamples.OutPath + "VideoTrimming-out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
        }
    }
}