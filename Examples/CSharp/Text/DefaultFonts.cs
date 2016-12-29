using System.IO;

using Aspose.Slides;
using System.Drawing.Imaging;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Text
{
    public class DefaultFonts
    {
        public static void Run()
        {
            // ExStart:DefaultFonts
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Text();

            // Use load options to define the default regualr and asian fonts// Use load options to define the default regualr and asian fonts
            LoadOptions loadOptions = new LoadOptions(LoadFormat.Auto);
            loadOptions.DefaultRegularFont = "Wingdings";
            loadOptions.DefaultAsianFont = "Wingdings";

            // Load the presentation
            using (Presentation pptx = new Presentation(dataDir + "DefaultFonts.pptx", loadOptions))
            {
                // Generate slide thumbnail
                pptx.Slides[0].GetThumbnail(1, 1).Save(dataDir + "output_out.png", ImageFormat.Png);

                // Generate PDF
                pptx.Save(dataDir + "output_out.pdf", SaveFormat.Pdf);

                // Generate XPS
                pptx.Save(dataDir + "output_out.xps", SaveFormat.Xps);
            }
            // ExEnd:DefaultFonts
        }
    }
}