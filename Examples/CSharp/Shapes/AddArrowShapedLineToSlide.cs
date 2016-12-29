using System.IO;
using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AddArrowShapedLineToSlide
    {
        public static void Run()
        {
            //ExStart:AddArrowShapedLineToSlide
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate PresentationEx class that represents the PPTX file
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add an autoshape of type line
                IAutoShape shp = sld.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0);

                // Apply some formatting on the line
                shp.LineFormat.Style = LineStyle.ThickBetweenThin;
                shp.LineFormat.Width = 10;

                shp.LineFormat.DashStyle = LineDashStyle.DashDot;

                shp.LineFormat.BeginArrowheadLength = LineArrowheadLength.Short;
                shp.LineFormat.BeginArrowheadStyle = LineArrowheadStyle.Oval;

                shp.LineFormat.EndArrowheadLength = LineArrowheadLength.Long;
                shp.LineFormat.EndArrowheadStyle = LineArrowheadStyle.Triangle;

                shp.LineFormat.FillFormat.FillType = FillType.Solid;
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Maroon;

                //Write the PPTX to Disk
                pres.Save(dataDir + "LineShape2_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddArrowShapedLineToSlide
        }
    }
}