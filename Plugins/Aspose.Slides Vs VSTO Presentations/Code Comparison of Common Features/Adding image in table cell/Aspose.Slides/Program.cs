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
            string FileName = @"E:\Aspose\Aspose Vs VSTO\Aspose.Slides Vs VSTO Presentations v 1.1\Sample Files\Removing Row Or Column in Table.pptx";
            string ImageFile = @"E:\Aspose\Aspose Vs VSTO\Aspose.Slides Vs VSTO Presentations v 1.1\Sample Files\AsposeLogo.jpg";
            
            Presentation MyPresentation = new Presentation(FileName);

            //Get First Slide
            ISlide sld = MyPresentation.Slides[0];

            //Creating a Bitmap Image object to hold the image file
            System.Drawing.Bitmap image = new Bitmap(ImageFile);

            //Create an IPPImage object using the bitmap object
            IPPImage imgx1 = MyPresentation.Images.AddImage(image);

            foreach (IShape shp in sld.Shapes)
                if (shp is ITable)
                {
                    ITable tbl = (ITable)shp;
                    
                    //Add image to first table cell
                    tbl[0, 0].FillFormat.FillType = FillType.Picture;
                    tbl[0, 0].FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
                    tbl[0, 0].FillFormat.PictureFillFormat.Picture.Image = imgx1;
                }
            //Save PPTX to Disk
            MyPresentation.Save(FileName, Export.SaveFormat.Pptx);
        }
    }
}
