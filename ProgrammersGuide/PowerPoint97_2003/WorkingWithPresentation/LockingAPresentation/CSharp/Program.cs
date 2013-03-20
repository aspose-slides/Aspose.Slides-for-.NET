//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace LockingAPresentation
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate presentation class
            Presentation pres = new Presentation(dataDir + "demo.ppt");

            //Loop through all the slides in the presentation
            foreach (Slide sld in pres.Slides)
                //Loop through all the shapes in the slide
                foreach (Shape shp in sld.Shapes)
                    //Lock each shape to be protected against the select
                    shp.Protection = ShapeProtection.LockSelect;

            //Write the presentation to disk
            pres.Write(dataDir + "demoLock.ppt");
        }
    }
}