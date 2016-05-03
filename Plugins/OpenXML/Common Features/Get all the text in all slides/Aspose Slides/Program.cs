// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string file = "Get all the text in a slide.pptx";
            int numberOfSlides = CountSlides(file);
            System.Console.WriteLine("Number of slides = {0}", numberOfSlides);
            string slideText;
            for (int i = 0; i < numberOfSlides; i++)
            {
                slideText = GetSlideText(file, i);
                System.Console.WriteLine("Slide #{0} contains: {1}", i + 1, slideText);
            }
            System.Console.ReadKey();
        }
        public static int CountSlides(string presentationFile)
        {
            //Instantiate PresentationEx class that represents PPTX
            using (PresentationEx pres = new PresentationEx(presentationFile))
            {
                return pres.Slides.Count;
            }
        }
        public static string GetSlideText(string docName, int index)
        {
            string sldText = "";
            //Instantiate PresentationEx class that represents PPTX
            using (PresentationEx pres = new PresentationEx(docName))
            {
                //Access the slide
                SlideEx sld = pres.Slides[index];

                //Iterate through shapes to find the placeholder
                foreach (ShapeEx shp in sld.Shapes)
                    if (shp.Placeholder != null)
                    {
                        //get the text of each placeholder
                        sldText += ((AutoShapeEx)shp).TextFrame.Text;
                    }

            }
            return sldText;
        }
    }
}
