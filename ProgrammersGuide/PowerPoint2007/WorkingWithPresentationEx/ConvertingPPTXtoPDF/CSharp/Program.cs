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

namespace ConvertingPPTXtoPDF
{
    public class Program
    {
        public static void Main(string[] args)
        {
            // The path to the documents directory.
            string dataDir = Path.GetFullPath("../../../Data/");

            // 1.
            // Conversion of PDF using default options.

            //Instantiate a PresentationEx object that represents a PPTX file
            PresentationEx pres = new PresentationEx(dataDir + "demo.pptx");

            //Saving the PPTX presentation to PDF document
            pres.Save(dataDir + "demo1.pdf", Aspose.Slides.Export.SaveFormat.Pdf);

            // Display result of conversion.
            System.Console.WriteLine("Conversion to PDF performed successfully with default options!");

            // 2.
            // Conversion to PDF using custom options.

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

            // Display result of conversion.
            System.Console.WriteLine("Conversion to PDF performed successfully with custom options!");
        }
    }
}