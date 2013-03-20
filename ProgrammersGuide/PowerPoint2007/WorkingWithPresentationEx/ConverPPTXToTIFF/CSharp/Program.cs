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
using System.Drawing;
using Aspose.Slides.Export;

namespace ConverPPTXToTIFF
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            //Instantiate a Presentation object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            // 1.
            //Saving the presentation to TIFF document
            pres.Save(dataDir + "demo.tiff", Aspose.Slides.Export.SaveFormat.Tiff);

            

            // 2.
            // Save as variable size TIFF
            //Instantiate the TiffOptions class
            Aspose.Slides.Export.TiffOptions opts = new Aspose.Slides.Export.TiffOptions();

            //Setting compression type
            opts.CompressionType = TiffCompressionTypes.Default;

            //Compression Types

            //Default - Specifies the default compression scheme (LZW).
            //None - Specifies no compression.
            //CCITT3
            //CCITT4
            //LZW
            //RLE

            //Depth – depends on the compression type and cannot be set manually.
            //Resolution unit – is always equal to “2” (dots per inch)

            //Setting image DPI
            opts.DpiX = 200;
            opts.DpiY = 100;

            //Set Image Size
            opts.ImageSize = new Size(1728, 1078);

            //Save the presentation to TIFF with specified image size
            pres.Save(dataDir + "demoCustomSize.tiff", Aspose.Slides.Export.SaveFormat.Tiff, opts);
        }
    }
}