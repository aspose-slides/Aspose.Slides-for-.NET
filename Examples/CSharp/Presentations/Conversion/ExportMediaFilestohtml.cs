using System.IO;
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
    class ExportMediaFilestohtml
    {
        public static void Run()
        {
            //ExStart:ExportMediaFilestohtml
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            // Loading a presentation
            using (Presentation pres = new Presentation(dataDir + "Media File.pptx"))
            {
                string path = dataDir;
                const string fileName = "ExportMediaFiles_out.html";
                const string baseUri = "http://www.example.com/";

                VideoPlayerHtmlController controller = new VideoPlayerHtmlController(path, fileName, baseUri);

                // Setting HTML options
                HtmlOptions htmlOptions = new HtmlOptions(controller);
                SVGOptions svgOptions = new SVGOptions(controller);

                htmlOptions.HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller);
                htmlOptions.SlideImageFormat = SlideImageFormat.Svg(svgOptions);

                // Saving the file
                pres.Save(Path.Combine(path, fileName), SaveFormat.Html, htmlOptions);
            }
            //ExEnd:ExportMediaFilestohtml
        }
    }
}
