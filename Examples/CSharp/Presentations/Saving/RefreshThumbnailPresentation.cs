using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following example shows how to save the PPTX presentation without generation of the new thumbnail.
*/

namespace CSharp.Presentations.Saving
{
    class RefreshThumbnailPresentation
    {
        public static void Run()
        {
            //Path for source presentation
            string pptxFile = Path.Combine(RunExamples.GetDataDir_PresentationSaving(), "Image.pptx");
            //Out path
            string resultPath = Path.Combine(RunExamples.OutPath, "result_with_old_thumbnail.pptx");

            using (Presentation pres = new Presentation(pptxFile))
            {
                //Remove all shapes from the slide
                pres.Slides[0].Shapes.Clear();

                //Save presentation
                pres.Save(resultPath, SaveFormat.Pptx, new PptxOptions()
                {
                    RefreshThumbnail = false
                });
            }
        }
    }
}
