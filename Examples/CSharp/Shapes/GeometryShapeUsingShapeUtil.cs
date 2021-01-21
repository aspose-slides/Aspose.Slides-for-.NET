
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides.Util;

/*
The example demonstrates using of ShapeUtil for editing shape geometry as System.Drawing.Drawing2D.GrpahicsPath object.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class GeometryShapeUsingShapeUtil
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "GeometryShapeUsingShapeUtil.pptx");

            using (Presentation pres = new Presentation())
            {
                // Create new shape
                GeometryShape shape = (GeometryShape)pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 100);

                // Get geometry path of the shape
                IGeometryPath originalPath = shape.GetGeometryPaths()[0];
                originalPath.FillMode = PathFillModeType.None;

                // Create new graphics path with text
                GraphicsPath graphicsPath = new GraphicsPath();
                graphicsPath.AddString("Text in shape", new FontFamily("Arial"), 1, 40, new PointF(10, 10), StringFormat.GenericDefault);

                // Convert graphics path to geometry path
                IGeometryPath textPath = ShapeUtil.GraphicsPathToGeometryPath(graphicsPath);
                textPath.FillMode = PathFillModeType.Normal;

                // Set combination of new geometry path and origin geometry path to the shape
                shape.SetGeometryPaths(new[] { originalPath, textPath });

                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
