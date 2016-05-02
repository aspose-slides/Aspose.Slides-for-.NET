// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System.Drawing;
using Aspose.Slides.Pptx;

namespace Slide_Thumbnail_to_JPEG
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation class that represents the presentation file
            using (PresentationEx pres = new PresentationEx(MyDir + "Slides Test Presentation.pptx"))
            {

                //Access the first slide
                SlideEx sld = pres.Slides[0];

                //Create a full scale image
                Bitmap bmp = sld.GetThumbnail(1f, 1f);

                //Save the image to disk in JPEG format
                bmp.Save(MyDir + "Test Thumbnail.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            }
        }
    }
}
