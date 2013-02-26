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

namespace AddPictureFrameToSlideEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PrseetationEx class that represents the PPTX
            PresentationEx pres = new PresentationEx();

            //Get the first slide
            SlideEx sld = pres.Slides[0];

            //Instantiate the ImageEx class
            System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir + "asp.jpg");
            ImageEx imgx = pres.Images.AddImage(img);

            //Add Picture Frame with height and width equivalent of Picture
            sld.Shapes.AddPictureFrame(ShapeTypeEx.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx);

            //Write the PPTX file to disk
            pres.Write(dataDir + "RectPicFrame.pptx");

            // Display Status.
            System.Console.WriteLine("Picture Frame added successfully.");
        }
    }
}