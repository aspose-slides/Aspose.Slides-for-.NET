//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Pptx;

namespace ReplacingTextInPlaceholder
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PresentationEx class that represents PPTX
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");


            //Access first slide
            SlideEx sld = pres.Slides[0];

            //Iterate through shapes to find the placeholder
            foreach (ShapeEx shp in sld.Shapes)
                if (shp.Placeholder != null)
                {
                    //Change the text of each placeholder
                    ((AutoShapeEx)shp).TextFrame.Text = "This is Placeholder";
                }

            //Write the PPTX to Disk
            pres.Write(dataDir + "output.pptx");
        }
    }
}