using System.IO;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Shapes
{
    public class AccessOLEObjectFrame
    {
        public static void Run()
        {
            //ExStart:AccessOLEObjectFrame
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Load the PPTX to Presentation object
            using (Presentation pres = new Presentation(dataDir + "AccessingOLEObjectFrame.pptx"))
            {
                // Access the first slide
                ISlide sld = pres.Slides[0];

                // Cast the shape to OleObjectFrame
                OleObjectFrame oleObjectFrame = sld.Shapes[0] as OleObjectFrame;

                // Read the OLE Object and write it to disk
                if (oleObjectFrame != null)
                {
                    // Get embedded file data
                    byte[] data = oleObjectFrame.EmbeddedFileData;

                    // Get embedded file extention
                    string fileExtention = oleObjectFrame.EmbeddedFileExtension;

                    // Create path for saving the extracted file
                    string extractedPath = dataDir + "excelFromOLE_out" + fileExtention;

                    // Save extracted data
                    using (FileStream fstr = new FileStream(extractedPath, FileMode.Create, FileAccess.Write))
                    {
                        fstr.Write(data, 0, data.Length);
                    }
                }
            }

            //ExEnd:AccessOLEObjectFrame
        }
    }
}