using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Shapes
{
    class ExtractEmbeddedFileDataFromOLEObject
    {
        public static void Run() {

            //ExStart:ExtractEmbeddedFileDataFromOLEObject

            // The documents directory path.
            string dataDir = RunExamples.GetDataDir_Shapes();

            string pptxFileName = dataDir +"TestOlePresentation.pptx";
            using (Presentation pres = new Presentation(pptxFileName))
            {
                int objectnum = 0;
                foreach (ISlide sld in pres.Slides)
                {
                    foreach (IShape shape in sld.Shapes)
                    {
                        if (shape is OleObjectFrame)
                        {
                            objectnum++;
                            OleObjectFrame oleFrame = shape as OleObjectFrame;
                            byte[] data = oleFrame.EmbeddedFileData;
                            string fileExtention = oleFrame.EmbeddedFileExtension;

                            string extractedPath = dataDir +"ExtractedObject_out" + objectnum + fileExtention;
                            using (FileStream fs = new FileStream(extractedPath, FileMode.Create))
                            {
                                fs.Write(data, 0, data.Length);
                            }
                        }
                    }
                }
            }

            //ExEnd:ExtractEmbeddedFileDataFromOLEObject

        }
    }
}
