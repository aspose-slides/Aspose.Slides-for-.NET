// Copyright (c) Aspose 2002-2014. All Rights Reserved.

using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Conversion.pptx";
            string destFileName = FilePath + "Converting to Tiff as defined format.tiff";
            
            //Instantiate a Presentation object that represents a Presentation file
            Presentation pres = new Presentation(srcFileName);
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
            pres.Save(destFileName, Aspose.Slides.Export.SaveFormat.Tiff, opts);
        }
    }
}
