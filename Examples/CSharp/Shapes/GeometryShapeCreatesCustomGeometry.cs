
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;

/*
The example demonstrates creation a shape with completely custom geometry.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class GeometryShapeCreatesCustomGeometry
    {
        public static void Run()
        {
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "GeometryShapeCreatesCustomGeometry.pptx");

            float R = 100, r = 50; // Outer and inner star radius

            // Create star geometry path
            GeometryPath starPath = CreateStarGeometry(R, r);

            using (Presentation pres = new Presentation())
            {
                // Create new shape
                GeometryShape shape = (GeometryShape)pres.Slides[0].Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, R * 2, R * 2);

                // Set new geometry path to the shape
                shape.SetGeometryPath(starPath);

                // Save the presentation
                pres.Save(resultPath, SaveFormat.Pptx);
            }
        }

        /// <summary>
        /// Creates star geometry path.
        /// </summary>
        /// <param name="outerRadius">Outet radius of a star figure.</param>
        /// <param name="innerRadiusr">inner radius of a star figure.</param>
        /// <returns>Geometry Path</returns>
        private static GeometryPath CreateStarGeometry(float outerRadius, float innerRadiusr)
        {
            GeometryPath starPath = new GeometryPath();
            List<PointF> points = new List<PointF>();

            int step = 72;

            for (int angle = -90; angle < 270; angle += step)
            {
                double radians = angle * (Math.PI / 180f);
                double x = outerRadius * Math.Cos(radians);
                double y = outerRadius * Math.Sin(radians);
                points.Add(new PointF((float)x + outerRadius, (float)y + outerRadius));

                radians = Math.PI * (angle + step / 2) / 180.0;
                x = innerRadiusr * Math.Cos(radians);
                y = innerRadiusr * Math.Sin(radians);
                points.Add(new PointF((float)x + outerRadius, (float)y + outerRadius));
            }

            starPath.MoveTo(points[0]);

            for (int i = 1; i < points.Count; i++)
            {
                starPath.LineTo(points[i]);
            }

            starPath.CloseFigure();

            return starPath;
        }
    }
}
