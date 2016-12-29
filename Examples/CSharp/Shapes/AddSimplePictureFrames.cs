using System.IO;
using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AddSimplePictureFrames
    {
        public static void Run()
        {
            //ExStart:AddSimplePictureFrames
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

                // Instantiate the ImageEx class
                System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir+ "aspose-logo.jpg");
                IPPImage imgx = pres.Images.AddImage(img);

                // Add Picture Frame with height and width equivalent of Picture
                sld.Shapes.AddPictureFrame(ShapeType.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx);

                //Write the PPTX file to disk
                pres.Save(dataDir + "RectPicFrame_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:AddSimplePictureFrames
        }
    }
}