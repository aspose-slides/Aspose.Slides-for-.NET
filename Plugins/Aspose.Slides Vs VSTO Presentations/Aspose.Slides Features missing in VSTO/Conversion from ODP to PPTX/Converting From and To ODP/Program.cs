using Aspose.Slides.Pptx;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Converting_From_and_To_ODP
{
    class Program
    {
        static void Main(string[] args)
        {
            ConvertedFromOdp();
            ConvertedToOdp();
        }
        public static void  ConvertedToOdp()
        {
            //Instantiate a Presentation object that represents a presentation file
            using (PresentationEx pres = new PresentationEx("ConversionFromPresentation.pptx"))
            {

                //Saving the PPTX presentation to PPTX format
                pres.Save("ConvertedToOdp", Aspose.Slides.Export.SaveFormat.Odp);
            }
        }
        public static void  ConvertedFromOdp()
        {
             //Instantiate a Presentation object that represents a presentation file
           using(PresentationEx pres = new PresentationEx("OpenOfficePresentation.odp"))
           {

               //Saving the PPTX presentation to PPTX format
              pres.Save("ConvertedFromOdp",Aspose.Slides.Export.SaveFormat.Pptx);
           }
        }
    }
}
