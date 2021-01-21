
using System.IO;
using Aspose.Slides.Export;

/*
This example demonstrates removing a segment from the existing geometry shape.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class GeometryShapeRemoveSegment
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "GeometryShapeRemoveSegment.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create new shape
                GeometryShape shape = (GeometryShape)pres.Slides[0].Shapes.AddAutoShape(ShapeType.Heart, 100, 100, 300, 300);

                // Get geometry path of the shape
                IGeometryPath path = shape.GetGeometryPaths()[0];

                // remove segment
                path.RemoveAt(2);

                // set new geometry path
                shape.SetGeometryPath(path);

                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
