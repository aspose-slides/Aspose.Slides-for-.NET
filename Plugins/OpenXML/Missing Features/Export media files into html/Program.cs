using Aspose.Slides;
using Aspose.Slides.Export;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, install it and then add its reference to this project. For any issues, questions or suggestions please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/
namespace Aspose.Plugins.AsposeVSOpenXML
{
    class Program
    {
        static void Main(string[] args)
        {
            string FilePath = @"..\..\..\Sample Files\";
            string srcFileName = FilePath + "Conversion.pptx";
            string destFileName =  "video.html";
            
            //Loading a presentation
            using (Presentation pres = new Presentation(srcFileName))
            {
                const string baseUri = "http://www.example.com/";

                VideoPlayerHtmlController controller = new VideoPlayerHtmlController(path: FilePath, fileName: destFileName, baseUri: baseUri);

                //Setting HTML options
                HtmlOptions htmlOptions = new HtmlOptions(controller);
                SVGOptions svgOptions = new SVGOptions(controller);

                htmlOptions.HtmlFormatter = HtmlFormatter.CreateCustomFormatter(controller);
                htmlOptions.SlideImageFormat = SlideImageFormat.Svg(svgOptions);

                //Saving the file
                pres.Save(destFileName, SaveFormat.Html, htmlOptions);
            }
        }
    }
}
