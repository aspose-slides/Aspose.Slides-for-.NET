// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            DeleteSlide("Delete a slide.pptx", 2);
        }
        
        public static void DeleteSlide(string presentationFile, int slideIndex)
        {
            //Instantiate a PresentationEx object that represents a PPTX file
            using (PresentationEx pres = new PresentationEx(presentationFile))
            {

                //Accessing a slide using its index in the slides collection
                SlideEx slide = pres.Slides[slideIndex];


                //Removing a slide using its reference
                pres.Slides.Remove(slide);


                //Writing the presentation as a PPTX file
                pres.Write(presentationFile);
            }
        }
    }
}
