using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Rendering_Printing
{
    class SupportOfDigitalSignature
    {

        public static void Run() {

            //ExStart:SupportOfDigitalSignature

            string dataDir = RunExamples.GetDataDir_Rendering();

            using (Presentation pres = new Presentation())
            {
                
                DigitalSignature signature = new DigitalSignature(dataDir + "testsignature1.pfx", @"testpass1");

                
                signature.Comments = "Aspose.Slides digital signing test.";

                
                pres.DigitalSignatures.Add(signature);

                
                pres.Save(dataDir + "SomePresentationSigned.pptx", SaveFormat.Pptx);
            }

            //ExEnd:SupportOfDigitalSignature



        }
    }
}
