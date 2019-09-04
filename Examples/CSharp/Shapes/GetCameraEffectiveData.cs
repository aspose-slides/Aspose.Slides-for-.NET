using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Shapes
{
    class GetCameraEffectiveData
    {
        public static void Run() {

            //ExStart:GetCameraEffectiveData

            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Shapes();

            using (Presentation pres = new Presentation(dataDir + "Presentation1.pptx"))
            {
                IThreeDFormatEffectiveData threeDEffectiveData = pres.Slides[0].Shapes[0].ThreeDFormat.GetEffective();

                Console.WriteLine("= Effective camera properties =");
                Console.WriteLine("Type: " + threeDEffectiveData.Camera.CameraType);
                Console.WriteLine("Field of view: " + threeDEffectiveData.Camera.FieldOfViewAngle);
                Console.WriteLine("Zoom: " + threeDEffectiveData.Camera.Zoom);

                
            }

            //ExEnd:GetCameraEffectiveData
        }
    }
}
