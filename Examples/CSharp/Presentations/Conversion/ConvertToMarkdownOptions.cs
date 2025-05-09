using System.IO;
using Aspose.Slides.Export;

/*
This example demonstrates how to save presentation to Markdown format using some convertion options.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertToMarkdownOptions
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "PresentationDemo.pptx");

            using (Presentation pres = new Presentation(presentationName))
            {
                // Path and folder name for saving markdown data
                string outPath = RunExamples.OutPath;

                MarkdownSaveOptions options = new MarkdownSaveOptions
                {
                    RemoveEmptyLines = true,
                    HandleRepeatedSpaces = HandleRepeatedSpaces.AlternateSpacesToNbsp,
                    SlideNumberFormat = "## Slide {0} -",
                    ShowSlideNumber = true,
                    ExportType = MarkdownExportType.TextOnly,
                    Flavor = Flavor.Default
                };

                // Save presentation in Markdown format
                pres.Save(Path.Combine(outPath, "pres-out.md"), SaveFormat.Md, options);
            }
        }
    }
}

