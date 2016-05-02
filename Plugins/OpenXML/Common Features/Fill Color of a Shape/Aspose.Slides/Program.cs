using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string docName = @"E:\Aspose\Aspose Vs OpenXML\Aspose.Slides Vs OpenXML Presentation v1.1\Sample Files\fill color of a shape.pptx";
            //Instantiate PrseetationEx class that represents the PPTX 
            using (Presentation pres = new Presentation())
            {
                //Get the first slide
                ISlide sld = pres.Slides[0];

                //Add autoshape of rectangle type
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150);

                //Set the fill type to Solid
                shp.FillFormat.FillType = FillType.Solid;

                //Set the color of the rectangle
                shp.FillFormat.SolidFillColor.Color = Color.Yellow;

                //Write the PPTX file to disk
                pres.Save(docName, SaveFormat.Pptx);
            }
        }
    }
}
