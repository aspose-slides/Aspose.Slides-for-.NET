using System.Drawing;
using Aspose.Slides.Export;
using Aspose.Slides.Pptx;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate the Presentation class that represents the presentation file
            string mypath = "";
            using (PresentationEx pres = new PresentationEx())
            {

                //Set the background color of the Master ISlide to Forest Green
                
                pres.Masters[0].Background.Type = BackgroundTypeEx.OwnBackground;
                pres.Masters[0].Background.FillFormat.FillType = FillTypeEx.Solid;
                pres.Masters[0].Background.FillFormat.SolidFillColor.Color = Color.ForestGreen;

                //Write the presentation to disk
                pres.Save(mypath + "Setting Background Color of Master Slide.pptx", SaveFormat.Pptx);

            }
 
        }
    }
}
