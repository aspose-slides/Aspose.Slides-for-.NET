using System.IO;
using Aspose.Slides;
using Aspose.Slides.Export;
using System.Drawing;

namespace Aspose.Slides.Examples.CSharp.Slides
{
    public class SetSlideBackgroundMaster
    {
        public static void Run()
        {
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Slides_Presentations();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            //Instantiate the Presentation class that represents the presentation file
            using (Presentation pres = new Presentation())
            {

                //Set the background color of the Master ISlide to Forest Green
                pres.Masters[0].Background.Type = BackgroundType.OwnBackground;
                pres.Masters[0].Background.FillFormat.FillType = FillType.Solid;
                pres.Masters[0].Background.FillFormat.SolidFillColor.Color = Color.ForestGreen;

                //Write the presentation to disk
                pres.Save(dataDir + "SetSlideBackgroundMaster.pptx", SaveFormat.Pptx);

            }
 
        }
    }
}