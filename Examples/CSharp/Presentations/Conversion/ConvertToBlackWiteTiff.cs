using System.IO;
using Aspose.Slides;
using Aspose.Slides.Animation;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example shows how to specify the algorithm for converting a color image to a black and white image.
*/

namespace Aspose.Slides.Examples.CSharp.Presentations.Conversion
{
    class ConvertToBlackWiteTiff
    {
        public static void Run()
        {
            // Path to source presentation
            string presentationName = Path.Combine(RunExamples.GetDataDir_Conversion(), "SimpleAnimations.pptx");
            // Path to output document
            string outFilePath = Path.Combine(RunExamples.OutPath, "BlackWhite_out.tiff");

            using (Presentation presentation = new Presentation(presentationName))
            {
                // Instantiate the TiffOptions class
                TiffOptions options = new TiffOptions()
                {
                    // Set compressio type
                    CompressionType = TiffCompressionTypes.CCITT4,
                    // Set convertion mode
                    BwConversionMode = BlackWhiteConversionMode.Dithering
                };

                // Save output file
                presentation.Save(outFilePath, new int[] { 2 }, SaveFormat.Tiff, options);
            }
        }
    }
}
