using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
This example demonstrates how to add the new vertical and horizontal drawing guides to a PowerPoint presentation
and change its color.
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


                // Getting the collection of the drawing guides for first master slide
                guides = pres.Masters[0].DrawingGuides;
                // Adding the new vertical drawing guide to the right of the slide center
                guides.Add(Orientation.Vertical, slideSize.Width / 2 + 20f);

                // Print the drawing guides of the first master slide
                Console.WriteLine(
                    string.Join(Environment.NewLine, guides.Select(g => $"{g.Orientation} {g.Position} {g.Color}")));

                // Change the color of the first drawing guide of the master slide
                guides[0].Color = Color.ForestGreen;

                // Print the drawing guides of the first master slide
                Console.WriteLine(
                    string.Join(Environment.NewLine, guides.Select(g => $"{g.Orientation} {g.Position} {g.Color}")));


                // Save presentation
                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
