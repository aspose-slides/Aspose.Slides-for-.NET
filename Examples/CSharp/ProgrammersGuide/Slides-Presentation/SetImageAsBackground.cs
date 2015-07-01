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
using Aspose.Slides.Export;

namespace CSharp.Slides
{
    public class SetImageAsBackground
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            //Instantiate the Presentation class that represents the presentation file

            using (Presentation pres = new Presentation(dataDir + "SetImageAsBackground.pptx"))
            {

                //Set the background with Image
                pres.Slides[0].Background.Type = BackgroundType.OwnBackground;
                pres.Slides[0].Background.FillFormat.FillType = FillType.Picture;
                pres.Slides[0].Background.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;

                //Set the picture
                System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir + "Tulips.jpg");

                //Add image to presentation's images collection
                IPPImage imgx = pres.Images.AddImage(img);

                pres.Slides[0].Background.FillFormat.PictureFillFormat.Picture.Image = imgx;

                //Write the presentation to disk
                pres.Save(dataDir + "ContentBG_Img.pptx", SaveFormat.Pptx);

            } 
        }
    }
}