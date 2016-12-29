using System.IO;

using Aspose.Slides;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class FormattedRectangle
    {
        public static void Run()
        {
            //ExStart:FormattedRectangle
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add autoshape of rectangle type
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 50);

                // Apply some formatting to rectangle shape
                shp.FillFormat.FillType = FillType.Solid;
                shp.FillFormat.SolidFillColor.Color = Color.Chocolate;

                // Apply some formatting to the line of rectangle
                shp.LineFormat.FillFormat.FillType = FillType.Solid;
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Black;
                shp.LineFormat.Width = 5;

                //Write the PPTX file to disk
                pres.Save(dataDir + "RectShp2_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            //ExEnd:FormattedRectangle
        }
    }
}