using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class SimpleRectangle
    {
        public static void Run()
        {
            //ExStart:SimpleRectangle
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add autoshape of rectangle type
                sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 50);

                //Write the PPTX file to disk
                pres.Save(dataDir+ "RectShp1_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SimpleRectangle
        }
    }
}