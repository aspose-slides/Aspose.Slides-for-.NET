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
    class ConvertIndividualSlide
    {
        //ExStart:ConvertIndividualSlide
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Conversion();
            using (Presentation presentation = new Presentation(dataDir + "Individual-Slide.pptx"))
            {
                HtmlOptions htmlOptions = new HtmlOptions();
                htmlOptions.HtmlFormatter = HtmlFormatter.CreateCustomFormatter(new CustomFormattingController());

                // Saving File              
                for (int i = 0; i < presentation.Slides.Count; i++)
                    presentation.Save(dataDir + "Individual Slide" + (i + 1) + "_out.html", new[] { i + 1 }, SaveFormat.Html, htmlOptions);
            }
        }

        public class CustomFormattingController : IHtmlFormattingController
        {
            void IHtmlFormattingController.WriteDocumentStart(IHtmlGenerator generator, IPresentation presentation)
            {}

            void IHtmlFormattingController.WriteDocumentEnd(IHtmlGenerator generator, IPresentation presentation)
            {}

            void IHtmlFormattingController.WriteSlideStart(IHtmlGenerator generator, ISlide slide)
            {
                generator.AddHtml(string.Format(SlideHeader, generator.SlideIndex + 1));
            }

            void IHtmlFormattingController.WriteSlideEnd(IHtmlGenerator generator, ISlide slide)
            {
                generator.AddHtml(SlideFooter);
            }

            void IHtmlFormattingController.WriteShapeStart(IHtmlGenerator generator, IShape shape)
            {}

            void IHtmlFormattingController.WriteShapeEnd(IHtmlGenerator generator, IShape shape)
            {}

            private const string SlideHeader = "<div class=\"slide\" name=\"slide\" id=\"slide{0}\">";
            private const string SlideFooter = "</div>";
        }
        //ExEnd:ConvertIndividualSlide
    }
}