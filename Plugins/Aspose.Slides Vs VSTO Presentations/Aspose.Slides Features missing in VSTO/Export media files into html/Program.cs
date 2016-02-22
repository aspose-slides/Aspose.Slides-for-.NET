using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Export_media_files_into_html
{
    class Program
    {
        static void Main(string[] args)
        {
            //Loading a presentation
            using (Presentation pres = new Presentation("example.pptx"))
            {
                const string path = "path";
                const string fileName = "video.html";
                const string baseUri = "http://www.example.com/";

                VideoPlayerHtmlController controller = new VideoPlayerHtmlController(path: path, fileName: fileName, baseUri: baseUri);

                //Setting HTML options
                HtmlOptions htmlOptions = new HtmlOptions(controller);
                SVGOptions svgOptions = new SVGOptions(controller);

                htmlOptions.HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller);
                htmlOptions.SlideImageFormat = SlideImageFormat.Svg(svgOptions);

                //Saving the file
                pres.Save(path + fileName, SaveFormat.Html, htmlOptions);
            }
        }
    }
}
