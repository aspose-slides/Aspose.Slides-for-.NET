using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Util;

/*
The following code shows how to modify a presentation and save it to a stream in its original format:
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Saving
{
    public class ToSaveFormatExample
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationSaving();
            string presentationPath = dataDir + "Presentation.pptm";

            using (var sourcePresentation = new Presentation(presentationPath))
            {
                // Modify the presentation as you need
                sourcePresentation.Slides.AddClone(sourcePresentation.Slides[0]);

                // Save the presentation to the stream in its original format
                using (var stream = new MemoryStream())
                    sourcePresentation.Save(stream, SlideUtil.ToSaveFormat(sourcePresentation.SourceFormat));
            }
        }
    }
}