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
            string ImageFilePath = @"E:\Aspose\Aspose Vs VSTO\Aspose.Slides Vs VSTO Presentations v 1.1\Sample Files\AddPicture.jpg";
            
            //Instantiate Prsentation class that represents the PPTX
            Presentation pres = new Presentation();

            //Get the first slide
            ISlide sld = pres.Slides[0];

            //Instantiate the ImageEx class
            Image img = (Image)new Bitmap(ImageFilePath);
            IPPImage imgx = pres.Images.AddImage(img);

            //Add Picture Frame with height and width equivalent of Picture
            sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx);
        }
    }
}
