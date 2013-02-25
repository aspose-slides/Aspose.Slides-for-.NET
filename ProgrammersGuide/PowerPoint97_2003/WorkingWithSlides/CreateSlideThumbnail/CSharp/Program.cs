//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;

using Aspose.Slides;

namespace CreateSlideThumbnail
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");

            //Accessing a slide using its slide position
            Slide slide = pres.GetSlideByPosition(1);

            //Getting the thumbnail image of the slide of a specified size
            Image image = slide.GetThumbnail(new Size(290, 230));

            //Saving the thumbnail image in jpeg format
            image.Save(dataDir + "thumbnail.jpg", ImageFormat.Jpeg);

            // Display Status.
            System.Console.WriteLine("Thumbnail created successfully.");
        }
    }
}