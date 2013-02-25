//////////////////////////////////////////////////////////////////////////
// Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
//
// This file is part of Aspose.Slides. The source code in this file
// is only intended as a supplement to the documentation, and is provided
// "as is", without warranty of any kind, either expressed or implied.
//////////////////////////////////////////////////////////////////////////
using System.IO;

using Aspose.Slides;

namespace ConvertingToPDFFile
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // 1.
            // Conversion using default options.

            //Instantiate a Presentation object that represents a PPT file
            Presentation pres = new Presentation(dataDir + "demo.ppt");

            //Saving the presentation to PDF document
            pres.Save(dataDir + "demo1.pdf", Aspose.Slides.Export.SaveFormat.Pdf);

            // Let user know about the conversion status.
            System.Console.WriteLine("Presentation saved to PDF with default options.");

            // 2. 
            // Conversion using custom options.

            //Instantiate the PdfOptions class
            Aspose.Slides.Export.PdfOptions opts = new Aspose.Slides.Export.PdfOptions();

            //Set Jpeg Quality
            opts.JpegQuality = 90;

            //Define behavior for meta files
            opts.SaveMetafilesAsPng = true;

            //Set Text Compression level
            opts.TextCompression = Aspose.Slides.Export.PdfTextCompression.Flate;

            //Define the PDF standard
            opts.Compliance = Aspose.Slides.Export.PdfCompliance.Pdf15;

            //Save the presentation to PDF with specified options
            pres.Save(dataDir + "demo2.pdf", Aspose.Slides.Export.SaveFormat.Pdf, opts);

            // Let user know about the conversion status.
            System.Console.WriteLine("Presentation saved to PDF with custom options.");
        }
    }
}