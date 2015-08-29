// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;

namespace OpenXML_Presentation
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Number of slides = {0}",
                CountSlides("Count the number of slides.pptx"));
            Console.ReadKey();
        }
        // Get the presentation object and pass it to the next CountSlides method.
        public static int CountSlides(string presentationFile)
        {
            // Open the presentation as read-only.
            using (PresentationDocument presentationDocument = PresentationDocument.Open(presentationFile, false))
            {
                // Pass the presentation to the next CountSlide method
                // and return the slide count.
                return CountSlides(presentationDocument);
            }
        }

        // Count the slides in the presentation.
        public static int CountSlides(PresentationDocument presentationDocument)
        {
            // Check for a null document object.
            if (presentationDocument == null)
            {
                throw new ArgumentNullException("presentationDocument");
            }

            int slidesCount = 0;

            // Get the presentation part of document.
            PresentationPart presentationPart = presentationDocument.PresentationPart;

            // Get the slide count from the SlideParts.
            if (presentationPart != null)
            {
                slidesCount = presentationPart.SlideParts.Count();
            }

            // Return the slide count to the previous method.
            return slidesCount;
        }
    }
}
