using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class RotatingShapes
    {
        public static void Run()
        {
            //ExStart:RotatingShapes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate PrseetationEx class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add autoshape of rectangle type
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150);

                // Rotate the shape to 90 degree
                shp.Rotation = 90;

                // Write the PPTX file to disk
                pres.Save(dataDir + "RectShpRot_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:RotatingShapes
        }
    }
}