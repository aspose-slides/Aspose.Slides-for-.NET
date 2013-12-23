//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AddEmbeddedFrame
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

            //Reading excel chart from the excel file and save as an array of bytes
            FileStream fstro = new FileStream(dataDir + "book1.xls", FileMode.Open, FileAccess.Read);
            byte[] b = new byte[fstro.Length];
            fstro.Read(b, 0, (int)fstro.Length);

            //Inserting the excel chart as new OleObjectFrame to a slide
            OleObjectFrame oof = slide.Shapes.AddOleObjectFrame(0, 0, pres.SlideSize.Width,
                           pres.SlideSize.Height, "Excel.Sheet.8", b);

            //Writing the presentation as a PPT file
            pres.Write(dataDir + "modified.ppt");

        }
    }
}