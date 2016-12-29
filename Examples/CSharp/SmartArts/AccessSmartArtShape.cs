using System.IO;

using Aspose.Slides;
using Aspose.Slides.SmartArt;

namespace Aspose.Slides.Examples.CSharp.SmartArts
{
    public class AccessSmartArtShape
    {
        public static void Run()
        {
            // ExStart:AccessSmartArtShape
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            // Load the desired the presentation
            using (Presentation pres = new Presentation(dataDir + "AccessSmartArtShape.pptx"))
            {

                // Traverse through every shape inside first slide
                foreach (IShape shape in pres.Slides[0].Shapes)
                {
                    // Check if shape is of SmartArt type
                    if (shape is ISmartArt)
                    {
                        // Typecast shape to SmartArtEx
                        ISmartArt smart = (ISmartArt)shape;
                        System.Console.WriteLine("Shape Name:" + smart.Name);

                    }
                }
            }
            // ExEnd:AccessSmartArtShape
        }
    }
}