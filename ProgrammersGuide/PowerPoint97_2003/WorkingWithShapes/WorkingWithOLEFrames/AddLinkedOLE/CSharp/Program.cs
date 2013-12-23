//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;
using System;

namespace AddLinkedOLE
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

            Presentation pres = new Presentation();

            //Access the first slide
            Slide slide = pres.GetSlideByPosition(1);

            // Adding substitute image image. This is mandatory operation.
            Picture pic = new Picture(pres, dataDir + "logo.jpg");
            pres.Pictures.Add(pic);

            // Creating new linked Ole object
            OleObjectFrame ole = slide.Shapes.AddOleObjectFrame(500, 100, 500, 500, "Excel.Sheet.8", new Guid("{00020820-0000-0000-C000-000000000046}"), null, dataDir + "book1.xls", null);
            ole.PictureId = pic.PictureId;

            // Replacing link and type of an object
            ole.SetObjectLink("Excel.Sheet.8", new Guid("{00020820-0000-0000-C000-000000000046}"), null, dataDir + "book1.xls", null);
            ole.PictureId = pic.PictureId;

            // Replacing link path without changing of object's type
            ole.SetObjectLink(null, dataDir + "book1.xls", null);

            //Save presentation
            pres.Write(dataDir + "modified.ppt");

        }
    }
}