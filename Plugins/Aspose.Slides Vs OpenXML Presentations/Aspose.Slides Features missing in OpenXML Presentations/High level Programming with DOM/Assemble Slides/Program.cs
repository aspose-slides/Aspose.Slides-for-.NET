// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Export;
using Aspose.Slides.Pptx;

namespace Assemble_Slides
{
    class Program
    {

        static void Main(string[] args)
        {
            AddingSlidetoPresentation();
            AccessingSlidesOfPresentation();
            RemovingSlides();
            ChangingPositionOfSlide();
        }
        public static void AddingSlidetoPresentation()
        {
            string MyDir = @"Files\";
            PresentationEx pres = new PresentationEx();

            //Instantiate SlideCollection class

            SlideExCollection slds = pres.Slides;
            for (int i = 0; i < pres.LayoutSlides.Count; i++)
            {
                //Add an empty slide to the Slides collection
                slds.AddEmptySlide(pres.LayoutSlides[i]);

            }

            //Save the PPTX file to the Disk
            pres.Write(MyDir + "EmptySlide.pptx");
        }
        public static void AccessingSlidesOfPresentation()
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Slides Test Presentation.pptx");
            //Accessing a slide using its slide index
            SlideEx slide = pres.Slides[0];

        }
        public static void RemovingSlides()
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Slides Test Presentation.pptx");

            //Accessing a slide using its index in the slides collection
            SlideEx slide = pres.Slides[0];
            //Removing a slide using its reference
            pres.Slides.Remove(slide);

            //Writing the presentation file
            pres.Write(MyDir + "modified.pptx");

        }
        public static void ChangingPositionOfSlide()
        {
            string MyDir = @"Files\";
            //Instantiate Presentation class to load the source presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Slides Test Presentation.pptx");
            {
                //Get the slide whose position is to be changed
                SlideEx sld = pres.Slides[0];
                //Set the new position for the slide
                sld.SlideNumber = 2;
                //Write the presentation to disk
                pres.Save(MyDir + "Changed Slide Position Presentation.pptx", SaveFormat.Pptx);
            }

        }
    }
}
