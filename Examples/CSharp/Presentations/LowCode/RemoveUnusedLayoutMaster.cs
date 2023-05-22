using System;
using System.Drawing;
using System.IO;
using Aspose.Slides.Export;
using Aspose.Slides;
using Aspose.Slides.Effects;

/*
This code demonstrates removing unused layout and master slides. 
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.LowCode
{
    class RemoveUnusedLayoutMaster
    {
        public static void Run()
        {
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_Slides_Presentations_LowCode(), "MultipleMaster.pptx");

            using (Presentation pres = new Presentation(pptxFileName))
            {
                Console.WriteLine("Master slides number in source presentation = " + pres.Masters.Count);
                Console.WriteLine("Layout slides number in source presentation = " + pres.LayoutSlides.Count);

                Aspose.Slides.LowCode.Compress.RemoveUnusedMasterSlides(pres);
                Aspose.Slides.LowCode.Compress.RemoveUnusedLayoutSlides(pres);

                Console.WriteLine("Master slides number in result presentation = " + pres.Masters.Count);
                Console.WriteLine("Layout slides number in result presentation = " + pres.LayoutSlides.Count);
            }
        }
    }
}

