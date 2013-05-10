//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace SettingBackgroundToPattern
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


            //Setting the fill type of the background to picture
            slide.Background.FillFormat.Type = FillType.Pattern;

                       
            //Setting Pattern Style
            slide.Background.FillFormat.PatternStyle = PatternStyle.DiagonalBrick;

            slide.Background.FillFormat.ForeColor = System.Drawing.Color.Chocolate;


            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");
        }
    }
}