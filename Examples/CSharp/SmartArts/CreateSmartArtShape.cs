using System.IO;
using Aspose.Slides;
using Aspose.Slides.SmartArt;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class CreateSmartArtShape
    {
        public static void Run()
        {
            // ExStart:CreateSmartArtShape
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);
            // Instantiate the presentation
            using (Presentation pres = new Presentation())
            {

                // Access the presentation slide
                ISlide slide = pres.Slides[0];

                // Add Smart Art Shape
                ISmartArt smart = slide.Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.BasicBlockList);

                // Saving presentation
                pres.Save(dataDir + "SimpleSmartArt_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:CreateSmartArtShape
        }
    }
}