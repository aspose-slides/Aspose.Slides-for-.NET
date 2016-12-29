using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AddPlainLineToSlide
    {
        public static void Run()
        {
            //ExStart:AddPlainLineToSlide
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate PresentationEx class that represents the PPTX file
            using (Presentation pres = new Presentation())
            {
                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add an autoshape of type line
                sld.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0);

                //Write the PPTX to Disk
                pres.Save(dataDir + "LineShape1_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddPlainLineToSlide
        }
    }
}