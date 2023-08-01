using System.IO;
using Aspose.Slides.DOM.Export.Markdown.SaveOptions;
using Aspose.Slides.Export;

/*
This example demonstrates how to save presentation to Markdown format.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertToMarkdown
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "PresentationDemo.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Path and folder name for saving markdown data
                string outPath = RunExamples.OutPath;

                // Create Markdown creation options
                MarkdownSaveOptions mdOptions = new MarkdownSaveOptions();
                // Set parameter for render all items (items that are grouped will be rendered together).
                mdOptions.ExportType = MarkdownExportType.Visual;
                // Set folder name for saving images
                mdOptions.ImagesSaveFolderName = "md-images";
                // Set path for folder images
                mdOptions.BasePath = outPath;

                // Save presentation in Markdown format
                pres.Save(Path.Combine(outPath, "pres.md"), SaveFormat.Md, mdOptions);
            }
        }
    }
}

