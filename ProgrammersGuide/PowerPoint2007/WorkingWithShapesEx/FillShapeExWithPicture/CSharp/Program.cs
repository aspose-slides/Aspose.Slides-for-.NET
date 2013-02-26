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

namespace FillShapeExWithPicture
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

            //Add auto shape of rectangle type
            int idx = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 50, 150, 75, 150);
            ShapeEx shp = sld.Shapes[idx];

            //Set the fill type to Picture
            shp.FillFormat.FillType = FillTypeEx.Picture;

            //Set the picture fill mode
            shp.FillFormat.PictureFillFormat.PictureFillMode = PictureFillModeEx.Tile;

            //Set the picture
            System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir + "asp.jpg");
            ImageEx imgx = pres.Images.AddImage(img);
            shp.FillFormat.PictureFillFormat.Picture.Image = imgx;

            //Write the PPTX file to disk
            pres.Write(dataDir + "RectShpPic.pptx");

            // Display Status.
            System.Console.WriteLine("Shape added and filled with image successfully.");
        }
    }
}