using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Shapes
{
    class GetShapeBevelEffectiveData
    {
        public static void Run()
        {

            //ExStart:GetShapeBevelEffectiveData

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            using (Presentation pres = new Presentation(dataDir + "Presentation1.pptx"))
            {
                IThreeDFormatEffectiveData threeDEffectiveData = pres.Slides[0].Shapes[0].ThreeDFormat.GetEffective();

                Console.WriteLine("= Effective shape's top face relief properties =");
                Console.WriteLine("Type: " + threeDEffectiveData.BevelTop.BevelType);
                Console.WriteLine("Width: " + threeDEffectiveData.BevelTop.Width);
                Console.WriteLine("Height: " + threeDEffectiveData.BevelTop.Height);


            }

            //ExEnd:GetShapeBevelEffectiveData
        }
    }
}
