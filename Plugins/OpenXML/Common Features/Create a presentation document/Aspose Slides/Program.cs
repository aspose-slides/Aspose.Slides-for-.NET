// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string filepath ="Create a presentation document.pptx";
            CreatePresentation(filepath);
        }
        public static void CreatePresentation(string filepath)
        {
            //Instantiate a Presentation object that represents a PPT file
            using (PresentationEx pres = new PresentationEx())
            {
                //Instantiate SlideExCollection calss
                SlideExCollection slds = pres.Slides;

                //Add an empty slide to the SlidesEx collection
                slds.AddEmptySlide(pres.LayoutSlides[0]);

                //Save your presentation to a file
                pres.Write(filepath);
            }
        }
    }
}
