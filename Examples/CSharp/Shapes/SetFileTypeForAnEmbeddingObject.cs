using Aspose.Slides;
using Aspose.Slides.DOM.Ole;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Shapes
{
    class SetFileTypeForAnEmbeddingObject
    {
        public static void Run() {

            //ExStart:SetFileTypeForAnEmbeddingObject

            using (Presentation pres = new Presentation())
            {
                // The path to the documents directory.
                string dataDir = RunExamples.GetDataDir_Shapes();

                // Add known Ole objects
                byte[] fileBytes = File.ReadAllBytes(dataDir + "test.zip");

                // Create Ole embedded file info
                IOleEmbeddedDataInfo dataInfo = new OleEmbeddedDataInfo(fileBytes, "zip");

                // Create OLE object
                IOleObjectFrame oleFrame = pres.Slides[0].Shapes.AddOleObjectFrame(150, 20, 50, 50, dataInfo);
                oleFrame.IsObjectIcon = true;


                pres.Save(dataDir + "SetFileTypeForAnEmbeddingObject.pptx", SaveFormat.Pptx);
            }

            //ExEnd:SetFileTypeForAnEmbeddingObject

        }
    }
}
