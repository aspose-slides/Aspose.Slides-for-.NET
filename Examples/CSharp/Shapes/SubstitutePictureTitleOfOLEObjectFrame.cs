using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Shapes
{
    class SubstitutePictureTitleOfOLEObjectFrame
    {
        public static void Run() {

            //ExStart:SubstitutePictureTitleOfOLEObjectFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();
            string oleSourceFile = dataDir +"ExcelObject.xlsx";
            string oleIconFile = dataDir + "Image.png";

            using (Presentation pres = new Presentation())
            {
                IPPImage image = null;
                ISlide slide = pres.Slides[0];

                // Add Ole objects
                byte[] allbytes = File.ReadAllBytes(oleSourceFile);
                IOleObjectFrame oof = slide.Shapes.AddOleObjectFrame(20, 20, 50, 50, "Excel.Sheet.12", allbytes);
                oof.IsObjectIcon = true;

                // Add image object
                byte[] imgBuf = File.ReadAllBytes(oleIconFile);
                using (MemoryStream ms = new MemoryStream(imgBuf))
                {
                    image = pres.Images.AddImage(new Bitmap(ms));
                }
                oof.SubstitutePictureFormat.Picture.Image = image;

                // Set caption to OLE icon
                oof.SubstitutePictureTitle = "Caption example";
            }

            //ExEnd:SubstitutePictureTitleOfOLEObjectFrame

        }
    }
}
