using System.IO;

using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class SimpleEllipse
    {
        public static void Run()
        {
            //ExStart:SimpleEllipse
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

                // Add autoshape of ellipse type
                sld.Shapes.AddAutoShape(ShapeType.Ellipse, 50, 150, 150, 50);

                //Write the PPTX file to disk
                pres.Save(dataDir + "EllipseShp1_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:SimpleEllipse
        }
    }
}