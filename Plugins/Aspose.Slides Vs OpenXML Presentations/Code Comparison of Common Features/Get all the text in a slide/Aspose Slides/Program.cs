// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Collections.Generic;
using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            foreach (string s in GetAllTextInSlide("Get all the text in a slide.pptx", 0))
                Console.WriteLine(s);
            Console.ReadKey();
        }
        // Get all the text in a slide.
        public static List<string> GetAllTextInSlide(string presentationFile, int slideIndex)
        {
            // Create a new linked list of strings.
            List<string> texts = new List<string>();

            //Instantiate PresentationEx class that represents PPTX
            using (PresentationEx pres = new PresentationEx(presentationFile))
            {

                //Access the slide
                SlideEx sld = pres.Slides[slideIndex];

                //Iterate through shapes to find the placeholder
                foreach (ShapeEx shp in sld.Shapes)
                    if (shp.Placeholder != null)
                    {
                        //get the text of each placeholder
                        texts.Add(((AutoShapeEx)shp).TextFrame.Text);
                    }

            }

            // Return an array of strings.
            return texts;
        }
    }
}
