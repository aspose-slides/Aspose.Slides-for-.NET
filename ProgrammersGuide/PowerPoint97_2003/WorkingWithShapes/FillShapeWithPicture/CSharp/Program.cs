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

namespace FillShapeWithPicture
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
            Slide slide = pres.GetSlideByPosition(2);

            //Adding an ellipse shape into the slide by defining its X,Y position, width
            //and height
            Shape shape = slide.Shapes.AddEllipse(1400, 1200, 3000, 2000);

            //Setting the fill type of the ellipse to picture
            shape.FillFormat.Type = FillType.Picture;

            //Creating a picture object that will be used to fill the ellipse
            Picture pic = new Picture(pres, dataDir + "demo.jpg");

            //Adding the picture object to pictures collection of the presentation
            //After the picture object is added, the picture is given a uniqe picture Id
            int picId = pres.Pictures.Add(pic);

            //Setting the picture Id of the shape fill to the Id of the picture object
            shape.FillFormat.PictureId = picId;

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");
        }
    }
}