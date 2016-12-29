using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class FillShapesPicture
    {
        public static void Run()
        {
            //ExStart:FillShapesPicture
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate PrseetationEx class that represents the PPTX
            using (Presentation pres = new Presentation())
            {

                // Get the first slide
                ISlide sld = pres.Slides[0];

                // Add autoshape of rectangle type
                IShape shp = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150);


                // Set the fill type to Picture
                shp.FillFormat.FillType = FillType.Picture;

                // Set the picture fill mode
                shp.FillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Tile;

                // Set the picture
                System.Drawing.Image img = (System.Drawing.Image)new Bitmap(dataDir + "Tulips.jpg");
                IPPImage imgx = pres.Images.AddImage(img);
                shp.FillFormat.PictureFillFormat.Picture.Image = imgx;

                //Write the PPTX file to disk
                pres.Save(dataDir + "RectShpPic_out.pptx", SaveFormat.Pptx);
                //ExEnd:FillShapesPicture
            }
        }
    }
}