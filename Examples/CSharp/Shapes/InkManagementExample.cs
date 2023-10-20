using System;
using System.Drawing;
using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.Ink;
using Microsoft.VisualStudio.TestTools.UnitTesting;

/*
This example shows how to get an Ink object and demonstrates how to change color and size of an Ink Brush.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class InkManagementExample
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Shapes(), "SimpleInk.pptx");
            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "SimpleInk_out.pptx");

            using (Presentation presentation = new Presentation(presentationName))
            {
                // Get Ink shape
                var inkShape = presentation.Slides[0].Shapes[0] as IInk;

                if (inkShape != null)
                {
                    Console.WriteLine("Width of the Ink shape = {0}", inkShape.Width);
                    Console.WriteLine("Height of the Ink shape = {0}", inkShape.Height);
                    Console.WriteLine("Brush height of the trace = {0}", inkShape.Traces[0].Brush.Size.Width);
                    Console.WriteLine("Brush color of the trace = {0}", inkShape.Traces[0].Brush.Color);

                    // Change color and size of the brush
                    inkShape.Traces[0].Brush.Color = Color.Red;
                    inkShape.Traces[0].Brush.Size = new SizeF(10f, 5f);
                }

                // Save presentation
                presentation.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
