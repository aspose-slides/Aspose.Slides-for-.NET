using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

/*
This example demonstrates how to get graphics path information from the ShapeElement.
*/

namespace CSharp.Shapes
{
    class ShapePathPointsExample
    {
        public static void Run()
        {
            // The documents directory path.
            string dataDir = RunExamples.GetDataDir_Shapes();
            string pptxFileName = dataDir + "PresetGeometry.pptx";

            using (Presentation pres = new Presentation(pptxFileName))
            {
                var autoShape = pres.Slides[0].Shapes[0] as AutoShape;
                
                IShapeElement[] elements = autoShape.CreateShapeElements();

                foreach (ShapeElement element in elements)
                {
                    Console.WriteLine("Start element");

                    byte[] types = element.PathTypes;
                    PointF[] points = element.PathPoints;
                    for (int i = 0; i < types.Length; i++)
                    {
                        switch (types[i])
                        {
                            case 0:
                                Console.WriteLine("Start point " + points[i].ToString());
                                break;
                            case 1:
                                Console.WriteLine("LineTo point " + points[i].ToString());
                                break;
                            case 3:
                                Console.WriteLine("Bezier spline point " + points[i].ToString());
                                break;
                            case 128:
                                Console.WriteLine("Close subpath point " + points[i].ToString());
                                break;
                            case 129:
                                Console.WriteLine("End point " + points[i].ToString());
                                break;
                        }
                    }
                }
            }
        }
    }
}