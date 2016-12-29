using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class FillShapesGradient
    {
        public static void Run()
        {
            //ExStart:FillShapesGradient
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Prseetation class that represents the PPTX// Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add autoshape of ellipse type
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Ellipse, 50, 150, 75, 150);

                // Apply some gradiant formatting to ellipse shape
                shp.FillFormat.FillType = FillType.Gradient;
                shp.FillFormat.GradientFormat.GradientShape = GradientShape.Linear;

                // Set the Gradient Direction
                shp.FillFormat.GradientFormat.GradientDirection = GradientDirection.FromCorner2;

                // Add two Gradiant Stops
                shp.FillFormat.GradientFormat.GradientStops.Add((float)1.0, PresetColor.Purple);
                shp.FillFormat.GradientFormat.GradientStops.Add((float)0, PresetColor.Red);

                //Write the PPTX file to disk
                pres.Save(dataDir + "EllipseShpGrad_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:FillShapesGradient
        }
    }
}