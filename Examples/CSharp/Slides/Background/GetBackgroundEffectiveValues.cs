using Aspose.Slides;
using Aspose.Slides.Examples.CSharp;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Slides.Background
{
    class GetBackgroundEffectiveValues
    {

        public static void Run()
        {
            //ExStart:GetBackgroundEffectiveValues
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations_Background();

            // Instantiate the Presentation class that represents the presentation file
            Presentation pres = new Presentation(dataDir + "SamplePresentation.pptx");

            IBackgroundEffectiveData effBackground = pres.Slides[0].Background.GetEffective();

            if (effBackground.FillFormat.FillType == FillType.Solid)
                Console.WriteLine("Fill color: " + effBackground.FillFormat.SolidFillColor);
            else
                Console.WriteLine("Fill type: " + effBackground.FillFormat.FillType);

            //ExEnd:GetBackgroundEffectiveValues
        }

    }
}
