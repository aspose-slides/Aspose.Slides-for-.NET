
using System.IO;

using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class FormatLines
    {
        public static void Run()
        {
            //ExStart:FormatLines
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
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 75);

                // Set the fill color of the rectangle shape
                shp.FillFormat.FillType = FillType.Solid;
                shp.FillFormat.SolidFillColor.Color = Color.White;

                // Apply some formatting on the line of the rectangle
                shp.LineFormat.Style = LineStyle.ThickThin;
                shp.LineFormat.Width = 7;
                shp.LineFormat.DashStyle = LineDashStyle.Dash;

                // Set the color of the line of rectangle
                shp.LineFormat.FillFormat.FillType = FillType.Solid;
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;

                //Write the PPTX file to disk
                pres.Save(dataDir + "RectShpLn_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:FormatLines
        }
    }
}