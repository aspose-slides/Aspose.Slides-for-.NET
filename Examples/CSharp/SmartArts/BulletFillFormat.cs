using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using Aspose.Slides.SmartArt;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.SmartArts
{
    class BulletFillFormat
    {
        public static void Run() {

            //ExStart:BulletFillFormat
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_SmartArts();

            using (Presentation presentation = new Presentation())
            {
                ISmartArt smart = presentation.Slides[0].Shapes.AddSmartArt(10, 10, 500, 400, SmartArtLayoutType.VerticalPictureList);
                ISmartArtNode node = smart.AllNodes[0];

                if (node.BulletFillFormat != null)
                {
                    Image img = (Image)new Bitmap(dataDir + "aspose-logo.jpg");
                    IPPImage image = presentation.Images.AddImage(img);
                    node.BulletFillFormat.FillType = FillType.Picture;
                    node.BulletFillFormat.PictureFillFormat.Picture.Image = image;
                    node.BulletFillFormat.PictureFillFormat.PictureFillMode = PictureFillMode.Stretch;
                }
                presentation.Save(dataDir +"out.pptx", SaveFormat.Pptx);
            }
            //ExEnd:BulletFillFormat
        }
    }
}
