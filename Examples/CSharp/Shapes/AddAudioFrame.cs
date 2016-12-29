using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AddAudioFrame
    {
        public static void Run()
        {
            //ExStart:AddAudioFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Load the wav sound file to stram
                FileStream fstr = new FileStream(dataDir+ "sampleaudio.wav", FileMode.Open, FileAccess.Read);

                // Add Audio Frame
                IAudioFrame af = sld.Shapes.AddAudioFrameEmbedded(50, 150, 100, 100, fstr);

                // Set Play Mode and Volume of the Audio
                af.PlayMode = AudioPlayModePreset.Auto;
                af.Volume = AudioVolumeMode.Loud;

                //Write the PPTX file to disk
                pres.Save(dataDir + "AudioFrameEmbed_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddAudioFrame
        }
    }
}