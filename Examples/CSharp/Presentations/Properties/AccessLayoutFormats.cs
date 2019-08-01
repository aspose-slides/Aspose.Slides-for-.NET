using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Properties
{
    class AccessLayoutFormats
    {
        public static void Run() {

            //ExStart:AccessLayoutFormats

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_PresentationProperties();

            using (Presentation pres = new Presentation(dataDir + "pres.pptx"))
            {
                foreach (ILayoutSlide layoutSlide in pres.LayoutSlides)
                {
                    IFillFormat[] fillFormats = layoutSlide.Shapes.Select(shape => shape.FillFormat).ToArray();
                    ILineFormat[] lineFormats = layoutSlide.Shapes.Select(shape => shape.LineFormat).ToArray();
                }
            }
            //ExEnd:AccessLayoutFormats

        }

    }
}
