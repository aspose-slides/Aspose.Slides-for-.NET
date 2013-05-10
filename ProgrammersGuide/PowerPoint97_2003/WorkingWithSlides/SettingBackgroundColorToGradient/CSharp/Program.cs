//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace SettingBackgroundColorToGradient
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


            //Disable following master background settings
            slide.FollowMasterBackground = false;


            //Setting the fill type of the background to gradient
            slide.Background.FillFormat.Type = FillType.Gradient;


            //Setting the background color to blue
            slide.Background.FillFormat.ForeColor = System.Drawing.Color.Blue;


            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");
        }
    }
}