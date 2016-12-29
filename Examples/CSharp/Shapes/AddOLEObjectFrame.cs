using System.IO;
using Aspose.Slides;
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
            Presentation pres = new Presentation();

            // Access the first slide
            ISlide sld = pres.Slides[0];

            // Load an cel file to stream
            FileStream fs = new FileStream(dataDir+ "book1.xlsx", FileMode.Open, FileAccess.Read);
            MemoryStream mstream = new MemoryStream();
            byte[] buf = new byte[4096];

            while (true)
            {
                int bytesRead = fs.Read(buf, 0, buf.Length);
                if (bytesRead <= 0)
                    break;
                mstream.Write(buf, 0, bytesRead);
            }

            // Add an Ole Object Frame shape
            IOleObjectFrame oof = sld.Shapes.AddOleObjectFrame(0, 0, pres.SlideSize.Size.Width, pres.SlideSize.Size.Height, "Excel.Sheet.12", mstream.ToArray());

            //Write the PPTX to disk
            pres.Save(dataDir + "OleEmbed_out.pptx", SaveFormat.Pptx);
            //ExEnd:AddOLEObjectFrame
        }
    }
}