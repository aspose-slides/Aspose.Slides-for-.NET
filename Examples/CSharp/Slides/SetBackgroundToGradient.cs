using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class SetBackgroundToGradient
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            // Instantiate the Presentation class that represents the presentation file
            using (Presentation pres = new Presentation(dataDir + "SetBackgroundToGradient.pptx"))
            {

                //Apply Gradiant effect to the Background
                pres.Slides[0].Background.Type = BackgroundType.OwnBackground;
                pres.Slides[0].Background.FillFormat.FillType = FillType.Gradient;
                pres.Slides[0].Background.FillFormat.GradientFormat.TileFlip = TileFlip.FlipBoth;

                //Write the presentation to disk
                pres.Save(dataDir + "ContentBG_Grad_out.pptx", SaveFormat.Pptx);
            }
 
        }
    }
}