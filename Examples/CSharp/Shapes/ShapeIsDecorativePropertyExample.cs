using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This sample demonstrates how to set the shape as “decorative” object (used for decorative purposes, for example, some stylistic objects).
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class ShapeIsDecorativePropertyExample
    {
        public static void Run()
        {
            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "DecorativeDemo.pptx");
            using (Presentation pres = new Presentation())
            {
                // Create new shape
                IShape shape1 = pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 10, 10, 100, 100);

                // Set shape as “decorative” object
                shape1.IsDecorative = true;

                // Save result
                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
