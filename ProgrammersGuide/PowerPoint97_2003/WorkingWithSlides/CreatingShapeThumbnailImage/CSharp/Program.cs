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
using System.Drawing.Imaging;

namespace CreatingShapeThumbnailImage
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


            //Iterate all shapes on a slide and create thumbnails
            ShapeCollection shapes = slide.Shapes;
            for (int i = 0; i < shapes.Count; i++)
            {
                Shape shape = shapes[i];


                //Getting the thumbnail image of the shape
                Image img = slide.GetThumbnail(new object[] { shape }, 1.0, 1.0, shape.ShapeRectangle);


                //Saving the thumbnail image in gif format
                img.Save(dataDir + "demo" + i + ".gif", ImageFormat.Gif);
            }
        }
    }
}