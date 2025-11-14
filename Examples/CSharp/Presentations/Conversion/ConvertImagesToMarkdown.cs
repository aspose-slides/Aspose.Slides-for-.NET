using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
This example shows how to control the process of saving and referencing images in the resulting Markdown file 
using the ImageSaving and SvgImageSaving events.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertImagesToMarkdown
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "demo_2.pptx");
            string outFilePath = Path.Combine(RunExamples.OutPath, "output_markdown.md");
            string imagesDir = Path.Combine(RunExamples.OutPath, "ExportedImages");

            // Check the catalog for images
            if (!Directory.Exists(imagesDir))
            {
                Directory.CreateDirectory(imagesDir);
            }

            var options = new MarkdownSaveOptions()
            {
                ImagesSaveFolderName = "Images",
                ExportType = MarkdownExportType.Visual
            };

            options.ImageSaving += (IImage image, ImageFormat format, ref string link) =>
            {
                //string imagesDir = "ExportedImages";
                format = ImageFormat.Jpeg; //Force output format to JPEG for all images.
                string fileName = "Image_" + Guid.NewGuid().ToString("N") + ".jpg";

                link = Path.Combine(imagesDir, fileName);
                image.Save(link, format);

                return true;
            };

            options.SvgImageSaving += (ISvgImage svgImage, ref string link) =>
            {
                //string imagesDir = "ExportedImages";
                string fileName = "Svg_" + Guid.NewGuid().ToString("N") + ".svg";

                link = Path.Combine(imagesDir, fileName);
                File.WriteAllBytes(link, svgImage.SvgData);

                return true;
            };

            using (var presentation = new Presentation(presentationName))
            {
                presentation.Save(outFilePath, SaveFormat.Md, options);
            }
        }
    }
}
