using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Media
{
    class ExtractAudio
    {
        public static void Run() {

            //ExStart:ExtractAudio

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Media();

            string presName = dataDir + "AudioSlide.ppt";
           
            // Instantiate Presentation class that represents the presentation file
            Presentation pres = new Presentation(presName);

            // Access the desired slide
            ISlide slide = pres.Slides[0];

            // Get the slideshow transition effects for slide
            ISlideShowTransition transition = slide.SlideShowTransition;

            //Extract sound in byte array
            byte[] audio = transition.Sound.BinaryData;

            System.Console.WriteLine("Length: " + audio.Length);
            //ExEnd:ExtractAudio

        }
    }
}
