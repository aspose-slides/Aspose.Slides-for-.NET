// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Number of slides = {0}",
                CountSlides("Count the number of slides.pptx"));
            Console.ReadKey();
        }
        public static int CountSlides(string presentationFile)
        {
            //Instantiate a PresentationEx object that represents a PPTX file
            using (PresentationEx pres = new PresentationEx(presentationFile))
            {

                return pres.Slides.Count;
            }
        }
    }
}
