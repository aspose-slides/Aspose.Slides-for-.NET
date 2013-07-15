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

namespace AccessingAnExistingTableInSlideEx
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate PresentationEx class that represents PPTX
            PresentationEx pres = new PresentationEx(dataDir + "table.pptx");

            //Access the first slide
            SlideEx sld = pres.Slides[0];

            //Initialize null TableEx
            TableEx tbl = null;

            //Iterate through the shapes and set a reference to the table found
            foreach (ShapeEx shp in sld.Shapes)
                if (shp is TableEx)
                    tbl = (TableEx)shp;

            //Set the text of the first column of second row
            tbl[0, 1].TextFrame.Text = "New";

            //Write the PPTX to Disk
            pres.Write(dataDir + "table_updated.pptx");

        }
    }
}