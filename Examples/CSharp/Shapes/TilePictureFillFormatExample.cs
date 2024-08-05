using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Export;

/*
The following example shows how to add the new Rectangle shape with a tiled picture fill and change the Tile properties.
*/

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    class TilePictureFillFormatExample
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // The path to output file
            string outFilePath = Path.Combine(RunExamples.OutPath, "ImageTileExample.pptx");

            using (Presentation pres = new Presentation())
            {
                ISlide firstSlide = pres.Slides[0];

                IPPImage ppImage;
                using (IImage newImage = Aspose.Slides.Images.FromFile(Path.Combine(dataDir, "image.png")))
                    ppImage = pres.Images.AddImage(newImage);

                // Adds the new Rectangle shape
                var newShape = firstSlide.Shapes.AddAutoShape(ShapeType.Rectangle, 0, 0, 350, 350);

                // Sets the fill type of the new shape to Picture
                newShape.FillFormat.FillType = FillType.Picture;

                // Sets the shape's fill image
                IPictureFillFormat pictureFillFormat = newShape.FillFormat.PictureFillFormat;
                pictureFillFormat.Picture.Image = ppImage;

                // Sets the picture fill mode to Tile and changes the properties
                pictureFillFormat.PictureFillMode = PictureFillMode.Tile;
                pictureFillFormat.TileOffsetX = -275;
                pictureFillFormat.TileOffsetY = -247;
                pictureFillFormat.TileScaleX = 120;
                pictureFillFormat.TileScaleY = 120;
                pictureFillFormat.TileAlignment = RectangleAlignment.BottomRight;
                pictureFillFormat.TileFlip = TileFlip.FlipBoth;

                pres.Save(outFilePath, SaveFormat.Pptx);
            }
        }
    }
}
