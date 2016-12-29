using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Conversion
{
    class ConvertToPDFWithHiddenSlides
    {
        public static void Run()
        {
            //ExStart:ConvertToPDFWithHiddenSlides
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();
            using (Presentation presentation = new Presentation(dataDir + "HiddingSlides.pptx"))
            {
                // Instantiate the PdfOptions class
                PdfOptions pdfOptions = new PdfOptions();

                // Specify that the generated document should include hidden slides
                pdfOptions.ShowHiddenSlides = true;

                // Save the presentation to PDF with specified options
                presentation.Save(dataDir + "PDFWithHiddenSlides_out.pdf", SaveFormat.Pdf, pdfOptions);
            }
            //ExEnd:ConvertToPDFWithHiddenSlides
        }
    }
}

