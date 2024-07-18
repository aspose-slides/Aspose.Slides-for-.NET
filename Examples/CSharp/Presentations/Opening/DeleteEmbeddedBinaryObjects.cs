using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.IO;

/*
This example shows how to remove embedded binary data from a presentation file while loading.
*/

namespace CSharp.Presentations.Opening
{
    class DeleteEmbeddedBinaryObjects
    {
        public static void Run()
        {
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_PresentationOpening(), "OlePptx.pptx");
            string outPath = Path.Combine(RunExamples.OutPath, "OlePptx-out.pptx");

            // Create loading options.
            LoadOptions loadOption = new LoadOptions
            {
                DeleteEmbeddedBinaryObjects = true
            };

            // Numbers of frames in a presentation.
            int emptyOleFrames;
            // Number of empty frames in a presentation.
            int oleFramesCount;

            using (Presentation pres = new Presentation(pptxFileName, loadOption))
            {
                oleFramesCount = GetOleObjectFrameCount(pres.Slides, out emptyOleFrames);
                Console.WriteLine("Number of OLE frames in source presentation = {0}", oleFramesCount);
                Console.WriteLine("Number of empty OLE frames in source presentation = {0}", emptyOleFrames);

                pres.Save(outPath, SaveFormat.Pptx);
                using (Presentation outPres = new Presentation(outPath))
                {
                    oleFramesCount = GetOleObjectFrameCount(outPres.Slides, out emptyOleFrames);
                    Console.WriteLine("Number of OLE frames in resulting presentation = {0}", oleFramesCount);
                    Console.WriteLine("Number of empty OLE frames in resulting presentation = {0}", emptyOleFrames);
                }
            }
        }

        private static int GetOleObjectFrameCount(ISlideCollection slides, out int emptyOleFrames)
        {
            int oleFramesCount = 0;

            emptyOleFrames = 0;
            foreach (ISlide sld in slides)
            {
                foreach (IShape shape in sld.Shapes)
                {
                    OleObjectFrame objectFrame = shape as OleObjectFrame;
                    if (objectFrame == null)
                        continue;

                    oleFramesCount++;

                    byte[] embeddedData = objectFrame.EmbeddedData.EmbeddedFileData;
                    if (embeddedData == null || embeddedData.Length == 0)
                        emptyOleFrames++;
                }
            }

            return oleFramesCount;
        }

    }
}
