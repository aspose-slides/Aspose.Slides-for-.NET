using System.IO;
using Aspose.Slides.Export;

/*
This example demonstrates using UseFrameSize and UseFrameRotation for rendering presentation shapes to SVG.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class SvgConvertionOptions
    {
        public static void Run()
        {
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "SvgShapesConvertion.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "SvgShapesConvertion.svg");

            using (Presentation presentation = new Presentation(presentationName))
            {
                // Create new SVG option
                SVGOptions svgOptions = new SVGOptions();

                // Set UseFrameSize property to include frame in a rendering area.
                svgOptions.UseFrameSize = true;

                // Set UseFrameRotation property to exclude rotation of the shape when rendering.
                svgOptions.UseFrameRotation = false;

                // Save a shape to svg using SvgOptions
                using (FileStream stream = new FileStream(outPath, FileMode.Create))
                {
                    presentation.Slides[0].Shapes[0].WriteAsSvg(stream, svgOptions);
                }
            }
        }
    }
}

