using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\..\Sample Files\";
            string FileName = FilePath + "Fill color of a shape.pptx";

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
                pres.Save(FileName, SaveFormat.Pptx);
            }
        }
    }
}
