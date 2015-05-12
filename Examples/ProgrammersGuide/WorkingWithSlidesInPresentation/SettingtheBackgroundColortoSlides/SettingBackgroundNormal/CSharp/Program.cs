//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;

namespace SettingBackgroundNormal
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate the Presentation class that represents the presentation file
            using (Presentation pres = new Presentation())
            {

                //Set the background color of the first ISlide to Blue
                pres.Slides[0].Background.Type = BackgroundType.OwnBackground;
                pres.Slides[0].Background.FillFormat.FillType = FillType.Solid;
                pres.Slides[0].Background.FillFormat.SolidFillColor.Color = Color.Blue;

                pres.Save(dataDir + "ContentBG.pptx", SaveFormat.Pptx);

            }
        }
    }
}