using System.IO;
using Aspose.Slides;
using System;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class FindShapeInSlide
    {
        //ExStart:FindShapeInSlide
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate a Presentation class that represents the presentation file
            using (Presentation p = new Presentation(dataDir + "FindingShapeInSlide.pptx"))
            {

                ISlide slide = p.Slides[0];
                // Alternative text of the shape to be found
                IShape shape = FindShape(slide, "Shape1");
                if (shape != null)
                {
                    Console.WriteLine("Shape Name: " + shape.Name);
                }
            }
        }
        
        // Method implementation to find a shape in a slide using its alternative text
        public static IShape FindShape(ISlide slide, string alttext)
        {
            // Iterating through all shapes inside the slide
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                // If the alternative text of the slide matches with the required one then
                // Return the shape
                if (slide.Shapes[i].AlternativeText.CompareTo(alttext) == 0)
                    return slide.Shapes[i];
            }
            return null;
        }
        //ExEnd:FindShapeInSlide
    }
}

