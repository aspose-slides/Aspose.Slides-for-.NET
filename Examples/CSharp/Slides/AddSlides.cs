using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class AddSlides
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Presentation class that represents the presentation file
            using (Presentation pres = new Presentation())
            {
                // Instantiate SlideCollection calss
                ISlideCollection slds = pres.Slides;

                for (int i = 0; i < pres.LayoutSlides.Count; i++)
                {
                    // Add an empty slide to the Slides collection
                    slds.AddEmptySlide(pres.LayoutSlides[i]);

                }

                // Save the PPTX file to the Disk
                pres.Save(dataDir + "EmptySlide_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

            }
        }
    }
}