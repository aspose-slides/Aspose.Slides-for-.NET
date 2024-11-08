using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
This example demonstrates how to add the new vertical and horizontal drawing guides to a PowerPoint presentation.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations
{
    public class GuidesProperties
    {
        public static void Run()
        {
            string outFilePath = Path.Combine(RunExamples.OutPath, "GuidesProperties-out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Getting slide size
                var slideSize = pres.SlideSize.Size;

                // Getting the collection of the drawing guides
                IDrawingGuidesCollection guides = pres.ViewProperties.SlideViewProperties.DrawingGuides;
                // Adding the new vertical drawing guide to the right of the slide center
                guides.Add(Orientation.Vertical, slideSize.Width / 2 + 12.5f);
                // Adding the new horizontal drawing guide below the slide center
                guides.Add(Orientation.Horizontal, slideSize.Height / 2 + 12.5f);

                // Save presentation
                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
