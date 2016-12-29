using System.IO;
using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class FormatJoinStyles
    {
        public static void Run()
        {
            //ExStart:FormatJoinStyles

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

                // Add three autoshapes of rectangle type
                IShape shp1 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 100, 150, 75);
                IShape shp2 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 300, 100, 150, 75);
                IShape shp3 = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 250, 150, 75);

                // Set the fill color of the rectangle shape
                shp1.FillFormat.FillType = FillType.Solid;
                shp1.FillFormat.SolidFillColor.Color = Color.Black;
                shp2.FillFormat.FillType = FillType.Solid;
                shp2.FillFormat.SolidFillColor.Color = Color.Black;
                shp3.FillFormat.FillType = FillType.Solid;
                shp3.FillFormat.SolidFillColor.Color = Color.Black;

                // Set the line width
                shp1.LineFormat.Width = 15;
                shp2.LineFormat.Width = 15;
                shp3.LineFormat.Width = 15;

                // Set the color of the line of rectangle
                shp1.LineFormat.FillFormat.FillType = FillType.Solid;
                shp1.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
                shp2.LineFormat.FillFormat.FillType = FillType.Solid;
                shp2.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;
                shp3.LineFormat.FillFormat.FillType = FillType.Solid;
                shp3.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue;

                // Set the Join Style
                shp1.LineFormat.JoinStyle = LineJoinStyle.Miter;
                shp2.LineFormat.JoinStyle = LineJoinStyle.Bevel;
                shp3.LineFormat.JoinStyle = LineJoinStyle.Round;

                // Add text to each rectangle
                ((IAutoShape)shp1).TextFrame.Text = "This is Miter Join Style";
                ((IAutoShape)shp2).TextFrame.Text = "This is Bevel Join Style";
                ((IAutoShape)shp3).TextFrame.Text = "This is Round Join Style";

                //Write the PPTX file to disk
                pres.Save(dataDir + "RectShpLnJoin_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:FormatJoinStyles
        }
    }
}