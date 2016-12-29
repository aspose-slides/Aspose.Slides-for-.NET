using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Slides.Layout
{
    class SetPDFPageSize
    {
        public static void Run()
        {
            //ExStart:SetPDFPageSize
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Layout();

            // ExStart:SetPDFPageSize
            // Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation();

            // Set SlideSize.Type Property 
            presentation.SlideSize.Type = SlideSizeType.A4Paper;

            // Set different properties of PDF Options
            PdfOptions opts = new  PdfOptions();
            opts.SufficientResolution = 600;

            // ExEnd:SetPDFPageSize
            // Save presentation to disk
            presentation.Save(dataDir + "SetPDFPageSize_out.pdf", SaveFormat.Pdf, opts);
            //ExEnd:SetPDFPageSize
        }
    }
}