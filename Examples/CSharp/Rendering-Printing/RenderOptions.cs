using Aspose.Slides.Export;
using Aspose.Slides;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace Aspose.Slides.Examples.CSharp.Rendering.Printing
{
    // This example demonstrates one of the possible use cases of IRenderingOptions interface
    //(getting slide thumbnails with different default font and slide's notes shown)

    class RenderOptions
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Rendering();
            string presPath = Path.Combine(dataDir, "RenderingOptions.pptx");

            using (Presentation pres = new Presentation(presPath))
            {
                IRenderingOptions renderingOpts = new RenderingOptions();
                NotesCommentsLayoutingOptions notesOptions = new NotesCommentsLayoutingOptions();
                notesOptions.NotesPosition = NotesPositions.BottomTruncated;
                renderingOpts.SlidesLayoutOptions = notesOptions;

                pres.Slides[0].GetThumbnail(renderingOpts, 4 / 3f, 4 / 3f).Save(Path.Combine(RunExamples.OutPath, "RenderingOptions-Slide1-Original.png"), ImageFormat.Png);

                renderingOpts.SlidesLayoutOptions = null;
                renderingOpts.DefaultRegularFont = "Arial Black";
                pres.Slides[0].GetThumbnail(renderingOpts, 4 / 3f, 4 / 3f).Save(Path.Combine(RunExamples.OutPath, "RenderingOptions-Slide1-ArialBlackDefault.png"), ImageFormat.Png);

                renderingOpts.DefaultRegularFont = "Arial Narrow";
                pres.Slides[0].GetThumbnail(renderingOpts, 4 / 3f, 4 / 3f).Save(Path.Combine(RunExamples.OutPath, "RenderingOptions-Slide1-ArialNarrowDefault.png"), ImageFormat.Png);
            }
        }
    }
}


