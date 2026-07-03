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
The code example below demonstrates how to configure image compression to convert images to HTML5.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class Html5PicturesCompressionExample
    {
        public static void Run()
        {
            // Path to source presentation
            string pptxFileName = Path.Combine(RunExamples.GetDataDir_Conversion(), "PresentationPic.pptx");

            // The path to output files
            string html5OutPath = Path.Combine(RunExamples.OutPath, "PresentationPic150.html");

            using (Presentation pres = new Presentation(pptxFileName))
            {
                // Set image compression level
                Html5Options options = new Html5Options()
                {
                    PicturesCompression = PicturesCompression.Dpi150
                };

                // Save result
                pres.Save(html5OutPath, SaveFormat.Html5, options);
            }
        }
    }
}