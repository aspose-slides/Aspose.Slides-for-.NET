using System.IO;
using Aspose.Slides.Export;

/*
The example demonstrates creation a composite custom shape from two GeometryPath objects.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class GeometryShapeCompositeObjects
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "GeometryShapeCompositeObjects.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create new shape
                GeometryShape shape = (GeometryShape)pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 200, 100);

                // Create first geometry path
                GeometryPath geometryPath0 = new GeometryPath();
                geometryPath0.MoveTo(0, 0);
                geometryPath0.LineTo(shape.Width, 0);
                geometryPath0.LineTo(shape.Width, shape.Height / 3);
                geometryPath0.LineTo(0, shape.Height / 3);
                geometryPath0.CloseFigure();

                // Create second geometry path
                GeometryPath geometryPath1 = new GeometryPath();
                geometryPath1.MoveTo(0, shape.Height / 3 * 2);
                geometryPath1.LineTo(shape.Width, shape.Height / 3 * 2);
                geometryPath1.LineTo(shape.Width, shape.Height);
                geometryPath1.LineTo(0, shape.Height);
                geometryPath1.CloseFigure();

                // Set shape geometry as composition of two geometry path
                shape.SetGeometryPaths(new GeometryPath[] { geometryPath0, geometryPath1 });

                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
