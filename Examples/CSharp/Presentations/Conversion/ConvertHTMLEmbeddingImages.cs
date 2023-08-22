using System.IO;
using Aspose.Slides.Export;

/*
When saving presentations in Html5, you can save images externally and the HTML document will use relative references to them. 
Next example demonstrates how to do that.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertHTMLEmbeddingImages
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "PresentationDemo.pptx");
            // Path to HTML document
            string outFilePath = Path.Combine(RunExamples.OutPath, "HTMLConvertion");

            using (Presentation pres = new Presentation(presentationName))
            {
                string outPath = RunExamples.OutPath; 

                Html5Options options = new Html5Options()
                {
                    // Force do not save images in HTML5 document
                    EmbedImages = false,
                    // Set path for external images
                    OutputPath = outPath
                };

                // Create directory for output HTML document
                if (!Directory.Exists(outFilePath))
                {
                    Directory.CreateDirectory(outFilePath);
                }

                // Save presentation in HTML5 format.
                pres.Save(Path.Combine(outFilePath, "pres.html"), SaveFormat.Html5, options);
            }
        }
    }
}

