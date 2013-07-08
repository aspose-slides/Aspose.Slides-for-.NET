//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;

namespace FindingShapeInSlideEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PresentationEx class that represents the PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            //Get the first slide
            SlideEx slide = pres.Slides[0];

            //Calling FindShape method and passing the slide reference with the
            //alternative text of the shape to be found
            ShapeEx shape = FindShape(slide, "Slides");

            if (shape != null)
            {
                System.Console.WriteLine("Shape Name: " + shape.Name);
                System.Console.WriteLine("Shape Height: " + shape.Height);
                System.Console.WriteLine("Shape Width: " + shape.Width);
            }
        }

        //Method implementation to find a shape in a slide using its alternative text
       public static ShapeEx FindShape(SlideEx slide, string alttext)
        {
            //Iterating through all shapes inside the slide
            for (int i = 0; i < slide.Shapes.Count; i++)
            {


                //If the alternative text of the slide matches with the required one then
                //return the shape
                if (slide.Shapes[i].AlternativeText.CompareTo(alttext) == 0)
                    return slide.Shapes[i];
            }
            return null;
        }

    }
}