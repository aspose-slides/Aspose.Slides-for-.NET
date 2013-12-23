//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace AccessingOLEFrame
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

            //Finding excel chart shape and obtaining its reference as OleObjectFrame
            OleObjectFrame oof = slide.Shapes[1] as OleObjectFrame;

            //Check, if OleObjectFrame is not null then read the object data of
            //OleObjectFrame as an array of bytes. Then write those bytes as an excel file
            if (oof != null)
            {
                FileStream fstr = new FileStream(dataDir + "book1.xls", FileMode.Create, FileAccess.Write);
                byte[] buf = oof.ObjectData;
                fstr.Write(buf, 0, buf.Length);
                fstr.Flush();
                fstr.Close();
                System.Console.WriteLine("Excel OLE Object written as excel1.xls file");
            }
 
        }
    }
}