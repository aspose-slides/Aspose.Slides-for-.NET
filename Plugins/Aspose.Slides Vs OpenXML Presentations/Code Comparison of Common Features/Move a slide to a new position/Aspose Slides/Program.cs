// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            MoveSlide("Move a slide to a new position.pptx", 0, 1);
        }
        // Move a slide to a different position in the slide order in the presentation.
        public static void MoveSlide(string presentationFile, int from, int to)
        {
            //Instantiate PresentationEx class to load the source PPTX file
            using (PresentationEx pres = new PresentationEx(presentationFile))
            {
                //Get the slide whose position is to be changed
                SlideEx sld = pres.Slides[from];
                SlideEx sld2 = pres.Slides[to];

                //Set the new position for the slide
                sld2.SlideNumber = from;
                sld.SlideNumber = to;

                //Write the PPTX to disk
                pres.Write(presentationFile);

            }
        }
    }
}
