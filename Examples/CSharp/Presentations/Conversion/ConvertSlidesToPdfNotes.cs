using System.Drawing;
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
    class ConvertSlidesToPdfNotes
    {
        public static void Run()
        {
            //ExStart:ConvertSlidesToPdfNotes
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            // Instantiate a Presentation object that represents a presentation file 
            Presentation presentation = new Presentation(dataDir + "SelectedSlides.pptx");
            Presentation auxPresentation = new Presentation();

            ISlide slide = presentation.Slides[0];

            auxPresentation.Slides.InsertClone(0, slide);

            // Setting Slide Type and Size 
            auxPresentation.SlideSize.Type = SlideSizeType.Custom;
            auxPresentation.SlideSize.Size = new SizeF(612F, 792F);
            auxPresentation.Save(dataDir + "PDFnotes_out.pdf", SaveFormat.PdfNotes);
            //ExEnd:ConvertSlidesToPdfNotes
        }
    }
}
