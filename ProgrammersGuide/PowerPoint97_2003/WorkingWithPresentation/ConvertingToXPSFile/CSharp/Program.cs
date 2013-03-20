//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ConvertingToXPSFile
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // 1.
            // Save presentation to XPS without using options provided by XpsOptions class.
            // Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");

            // Saving the presentation to TIFF document
            pres.Save(dataDir + "demo1.xps", Aspose.Slides.Export.SaveFormat.Xps);


            // 2.
            // Save presentation to XPS using options provided by XpsOptions class.
            //Instantiate the XpsOptions class
            Aspose.Slides.Export.XpsOptions opts = new Aspose.Slides.Export.XpsOptions();

            //Save MetaFiles as PNG
            opts.SaveMetafilesAsPng = true;

            //Save the presentation to XPS document
            pres.Save(dataDir + "demo2.xps", Aspose.Slides.Export.SaveFormat.Xps, opts);
        }
    }
}