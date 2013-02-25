//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;
using System.Drawing;

using Aspose.Slides;
using Aspose.Slides.Pptx;

namespace CreateSlideThumbnail
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a PresentationEx class that represents the PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            //Access the first slide
            SlideEx sld = pres.Slides[0];

            //Create a full scale image
            Bitmap bmp = sld.GetThumbnail(1f, 1f);

            //Save the image to disk in JPEG format
            bmp.Save(dataDir + "thumbnail.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            // Display Status.
            System.Console.WriteLine("Thumbnail created successfully.");
        }
    }
}