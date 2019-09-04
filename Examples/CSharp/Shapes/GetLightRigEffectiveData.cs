using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Shapes
{
    class GetLightRigEffectiveData
    {
        public static void Run()
        {

            //ExStart:GetLightRigEffectiveData

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            using (Presentation pres = new Presentation(dataDir + "Presentation1.pptx"))
            {
                IThreeDFormatEffectiveData threeDEffectiveData = pres.Slides[0].Shapes[0].ThreeDFormat.GetEffective();

                Console.WriteLine("= Effective light rig properties =");
                Console.WriteLine("Type: " + threeDEffectiveData.LightRig.LightType);
                Console.WriteLine("Direction: " + threeDEffectiveData.LightRig.Direction);


            }

            //ExEnd:GetLightRigEffectiveData
        }
    }
}
