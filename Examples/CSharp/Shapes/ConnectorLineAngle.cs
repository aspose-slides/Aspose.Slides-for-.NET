using System.IO;
using Aspose.Slides;
using System;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class ConnectorLineAngle
    {
        //ExStart:ConnectorLineAngle
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            Presentation pres = new Presentation(dataDir + "ConnectorLineAngle.pptx");
            Slide slide = (Slide)pres.Slides[0];
            Shape shape;
            for (int i = 0; i < slide.Shapes.Count; i++)
            {
                double dir = 0.0;
                shape = (Shape)slide.Shapes[i];
                if (shape is AutoShape)
                {
                    AutoShape ashp = (AutoShape)shape;
                    if (ashp.ShapeType == ShapeType.Line)
                    {
                        dir = getDirection(ashp.Width, ashp.Height, Convert.ToBoolean(ashp.Frame.FlipH), Convert.ToBoolean(ashp.Frame.FlipV));
                    }
                }
                else if (shape is Connector)
                {
                    Connector ashp = (Connector)shape;
                    dir = getDirection(ashp.Width, ashp.Height, Convert.ToBoolean(ashp.Frame.FlipH), Convert.ToBoolean(ashp.Frame.FlipV));
                }

                Console.WriteLine(dir);
            }

        }
        public static double getDirection(float w, float h, bool flipH, bool flipV)
        {
            float endLineX = w * (flipH ? -1 : 1);
            float endLineY = h * (flipV ? -1 : 1);
            float endYAxisX = 0;
            float endYAxisY = h;
            double angle = (Math.Atan2(endYAxisY, endYAxisX) - Math.Atan2(endLineY, endLineX));
            if (angle < 0) angle += 2 * Math.PI;
            return angle * 180.0 / Math.PI;
        }
        //ExEnd:ConnectorLineAngle
    }

}