using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to set visual rendering style of a two-color gradient.
*/
namespace CSharp.Slides.Views
{
    class GradientStyleRendering
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Slides_Views(), "GradientStyleExample.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "GradientStyleExample-out.png");

            using (Presentation pres = new Presentation(presentationName))
            {
                RenderingOptions options = new RenderingOptions();

                // Set rendering the two-color gradient according to its appearance in the PowerPoint user interface.
                options.GradientStyle = GradientStyle.PowerPointUI;

                // Get the image.
                IImage img = pres.Slides[0].GetImage(options, 2f, 2f);

                // Save image.
                img.Save(outPath, ImageFormat.Png);
            }
        }
    }
}
