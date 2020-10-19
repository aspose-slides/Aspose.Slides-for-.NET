using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes 
{
    public class EmbeddedVideoFrame
    {
        public static void Run()
        {
            //ExStart:EmbeddedVideoFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();
            string videoDir = RunExamples.GetDataDir_Video();
            string resultPath = Path.Combine(RunExamples.OutPath, "VideoFrame_out.pptx");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);
            // Instantiate Presentation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Embedd vide inside presentation
                IVideo vid = pres.Videos.AddVideo(new FileStream(videoDir + "Wildlife.mp4", FileMode.Open), LoadingStreamBehavior.ReadStreamAndRelease);

                // Add Video Frame
                IVideoFrame vf = sld.Shapes.AddVideoFrame(50, 150, 300, 350, vid);

                // Set video to Video Frame
                vf.EmbeddedVideo = vid;

                // Set Play Mode and Volume of the Video
                vf.PlayMode = VideoPlayModePreset.Auto;
                vf.Volume = AudioVolumeMode.Loud;

                // Write the PPTX file to disk
                pres.Save(resultPath, SaveFormat.Pptx);
            }
            //ExEnd:EmbeddedVideoFrame
        }
    }
}