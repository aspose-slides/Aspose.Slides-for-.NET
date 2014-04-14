//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System.Drawing;

namespace GeneratingThumbnailFromSlide
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "Aspose.pptx"))
            {

                //Access the first slide
                ISlide sld = pres.Slides[0];

                //Create a full scale image
                Bitmap bmp = sld.GetThumbnail(1f, 1f);

                //Save the image to disk in JPEG format
                bmp.Save(dataDir + "Thumbnail.jpg", System.Drawing.Imaging.ImageFormat.Jpeg);

            }
        }
    }
}