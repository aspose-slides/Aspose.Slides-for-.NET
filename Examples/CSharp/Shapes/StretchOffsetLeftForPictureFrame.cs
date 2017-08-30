using System.IO;
using Aspose.Slides;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class StretchOffsetLeftForPictureFrame
    {
        public static void Run()
        {
            //ExStart:StretchOffsetLeftForPictureFrame
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
                ISlide slide = pres.Slides[0];

                // Instantiate the ImageEx class
                System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir + "aspose-logo.jpg");
                IPPImage imgEx = pres.Images.AddImage(img);

                // Add an AutoShape of Rectangle type
                IAutoShape aShape = slide.Shapes.AddAutoShape(ShapeType.Rectangle, 100, 100, 300, 300);

                // Set shape's fill type
                aShape.FillFormat.FillType = FillType.Picture;

                // Set shape's picture fill mode
                aShape.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;

                // Set image to fill the shape
                aShape.FillFormat.PictureFillFormat.Picture.Image = imgEx;

                // Specify image offsets from the corresponding edge of the shape's bounding box
                aShape.FillFormat.PictureFillFormat.StretchOffsetLeft = 25;
                aShape.FillFormat.PictureFillFormat.StretchOffsetRight = 25;
                aShape.FillFormat.PictureFillFormat.StretchOffsetTop = -20;
                aShape.FillFormat.PictureFillFormat.StretchOffsetBottom = -10;


                //Write the PPTX file to disk
                pres.Save(dataDir + "StretchOffsetLeftForPictureFrame_out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:StretchOffsetLeftForPictureFrame
        }
    }
}
