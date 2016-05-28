using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace CSharp.ProgrammersGuide.Presentations
{
    class ExportMediaFilestohtml
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Presentations();

            // Loading a presentation
            using (Presentation pres = new Presentation(dataDir + "Media File.pptx"))
            {
                string path = dataDir;
                const string fileName = "ExportMediaFiles.html";
                const string baseUri = "http://www.example.com/";

                VideoPlayerHtmlController controller = new VideoPlayerHtmlController(path: path, fileName: fileName, baseUri: baseUri);

                // Setting HTML options
                HtmlOptions htmlOptions = new HtmlOptions(controller);
                SVGOptions svgOptions = new SVGOptions(controller);

                htmlOptions.HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller);
                htmlOptions.SlideImageFormat = SlideImageFormat.Svg(svgOptions);

                // Saving the file
                pres.Save(Path.Combine(path, fileName), SaveFormat.Html, htmlOptions);
            }
        }
    }
}
