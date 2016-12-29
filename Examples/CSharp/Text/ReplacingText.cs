using System.IO;

using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class ReplacingText
    {
        public static void Run()
        {
            // ExStart:ReplacingText
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Instantiate Presentation class that represents PPTX// Instantiate Presentation class that represents PPTX
            using (Presentation pres = new Presentation(dataDir + "ReplacingText.pptx"))
            {

                // Access first slide
                ISlide sld = pres.Slides[0];

                // Iterate through shapes to find the placeholder
                foreach (IShape shp in sld.Shapes)
                    if (shp.Placeholder != null)
                    {
                        // Change the text of each placeholder
                        ((IAutoShape)shp).TextFrame.Text = "This is Placeholder";
                    }

                // Save the PPTX to Disk
                pres.Save(dataDir + "output_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx);
            }
            // ExEnd:ReplacingText
        }
    }
}