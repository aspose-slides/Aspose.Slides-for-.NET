using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class CustomOptionsPDFConversion
    {
        public static void Run()
        {
            //ExStart:CustomOptionsPDFConversion
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            // Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(dataDir + "ConvertToPDF.pptx"))
            {
                // Instantiate the PdfOptions class
                PdfOptions pdfOptions = new PdfOptions();

                // Set Jpeg Quality
                pdfOptions.JpegQuality = 90;

                // Define behavior for metafiles
                pdfOptions.SaveMetafilesAsPng = true;

                // Set Text Compression level
                pdfOptions.TextCompression = PdfTextCompression.Flate;

                // Define the PDF standard
                pdfOptions.Compliance = PdfCompliance.Pdf15;

                // Save the presentation to PDF with specified options
                pres.Save(dataDir + "Custom_Option_Pdf_Conversion_out.pdf", SaveFormat.Pdf, pdfOptions);
            }
            //ExEnd:CustomOptionsPDFConversion
        }
    }
}