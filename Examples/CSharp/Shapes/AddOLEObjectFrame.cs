using System.IO;
using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Shapes 
{
    public class AddOLEObjectFrame
    {
        public static void Run()
        {
            //ExStart:AddOLEObjectFrame

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Prseetation class that represents the PPTX
            using (Presentation pres = new Presentation())
            {
                // Access the first slide
                ISlide sld = pres.Slides[0];

                // Load an cel file to stream
                MemoryStream mstream = new MemoryStream();
                using (FileStream fs = new FileStream(dataDir + "book1.xlsx", FileMode.Open, FileAccess.Read))
                {
                    byte[] buf = new byte[4096];

                    while (true)
                    {
                        int bytesRead = fs.Read(buf, 0, buf.Length);
                        if (bytesRead <= 0)
                            break;
                        mstream.Write(buf, 0, bytesRead);
                    }
                }

                // Create data object for embedding
                IOleEmbeddedDataInfo dataInfo = new OleEmbeddedDataInfo(mstream.ToArray(), "xlsx");

                // Add an Ole Object Frame shape
                IOleObjectFrame oleObjectFrame = sld.Shapes.AddOleObjectFrame(0, 0, pres.SlideSize.Size.Width,
                    pres.SlideSize.Size.Height, dataInfo);

                //Write the PPTX to disk
                pres.Save(dataDir + "OleEmbed_out.pptx", SaveFormat.Pptx);
            }

            //ExEnd:AddOLEObjectFrame
        }
    }
}