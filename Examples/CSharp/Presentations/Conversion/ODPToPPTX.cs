using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Presentations.Conversion
{
    class ODPToPPTX
    {

        public static void Run() {

            //ExStart:ODPToPPTX

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Conversion();

            
            string srcFileName = dataDir + "AccessOpenDoc.odp";
            string destFileName = dataDir + "AccessOpenDoc.pptx";
            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(srcFileName))
            {
                //Saving the PPTX presentation to PPTX format
                pres.Save(destFileName, SaveFormat.Pptx);
            }
            //ExEnd:ODPToPPTX



        }
    }
}
