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

namespace ConvertPPTXToXPS
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a PresentationEx object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");


            // 1.
            //Saving the presentation to TIFF document
            pres.Save(dataDir + "output.xps", Aspose.Slides.Export.SaveFormat.Xps);


            // 2.
            //Instantiate the TiffOptions class
            Aspose.Slides.Export.XpsOptions opts = new Aspose.Slides.Export.XpsOptions();

            //Save MetaFiles as PNG
            opts.SaveMetafilesAsPng = true;

            //Save the presentation to XPS document
            pres.Save(dataDir + "outputWithXPSOptions.xps", Aspose.Slides.Export.SaveFormat.Xps, opts);
        }
    }
}