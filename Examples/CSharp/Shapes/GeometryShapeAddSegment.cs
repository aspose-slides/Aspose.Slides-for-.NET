using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

/*
This example demonstrates adding new segment to the existing geometry shape.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class GeometryShapeAddSegment
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "GeometryShapeAddSegment.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create new shape
                GeometryShape shape = (GeometryShape)pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);
                // Get geometry path of the shape
                IGeometryPath geometryPath = shape.GetGeometryPaths()[0];

                // Add two lines to geometry path
                geometryPath.LineTo(100, 50, 1);
                geometryPath.LineTo(100, 50, 4);

                // Assign edited geometry path to the shape
                shape.SetGeometryPath(geometryPath);

                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
