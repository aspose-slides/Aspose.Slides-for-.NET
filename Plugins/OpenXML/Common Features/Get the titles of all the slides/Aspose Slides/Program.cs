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
            foreach (string s in GetSlideTitles("Get the titles of all the slides.pptx"))
                Console.WriteLine(s);
            Console.ReadKey();
        }
        // Get a list of the titles of all the slides in the presentation.
        public static IList<string> GetSlideTitles(string presentationFile)
        {
            // Create a new linked list of strings.
            List<string> texts = new List<string>();

            //Instantiate PresentationEx class that represents PPTX
            using (PresentationEx pres = new PresentationEx(presentationFile))
            {

                //Access all the slides
                foreach (SlideEx sld in pres.Slides)
                {

                    //Iterate through shapes to find the placeholder
                    foreach (ShapeEx shp in sld.Shapes)
                        if (shp.Placeholder != null)
                        {
                            if (IsTitleShape(shp))
                            {
                                //get the text of placeholder
                                texts.Add(((AutoShapeEx)shp).TextFrame.Text);
                            }
                        }
                }
            }

            // Return an array of strings.
            return texts;
        }
        // Determines whether the shape is a title shape.
        private static bool IsTitleShape(ShapeEx shape)
        {
            switch (shape.Placeholder.Type)
            {
                case PlaceholderTypeEx.Title:
                case PlaceholderTypeEx.CenteredTitle:
                    return true;
                default:
                    return false;
            }
        }
    }
}
