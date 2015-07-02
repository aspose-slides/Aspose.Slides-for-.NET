//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.Shapes
{
    public class AddAudioFrame
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                //Get the first slide
                ISlide sld = pres.Slides[0];

                //Load the wav sound file to stram
                FileStream fstr = new FileStream(dataDir+ "sampleaudio.wav", FileMode.Open, FileAccess.Read);

                //Add Audio Frame
                IAudioFrame af = sld.Shapes.AddAudioFrameEmbedded(50, 150, 100, 100, fstr);

                //Set Play Mode and Volume of the Audio
                af.PlayMode = AudioPlayModePreset.Auto;
                af.Volume = AudioVolumeMode.Loud;

                //Write the PPTX file to disk
                pres.Save(dataDir+ "AudioFrameEmbed.pptx", SaveFormat.Pptx);
            }

            
            
        }
    }
}