// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides.Pptx;

namespace Converting_to_Tiff
{
    class Program
    {
        static void Main(string[] args)
        {
            string MyDir = @"Files\";
            //Instantiate a Presentation object that represents a Presentation file
            PresentationEx pres = new PresentationEx(MyDir + "Conversion.ppt");
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
            pres.Save(MyDir + "Converted.tiff", Aspose.Slides.Export.SaveFormat.Tiff, opts);
        }
    }
}
