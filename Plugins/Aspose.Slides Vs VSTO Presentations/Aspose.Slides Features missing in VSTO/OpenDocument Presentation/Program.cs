using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenDocument_Presentation
{
    class Program
    {
        static void Main(string[] args)
        {
            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation("Source.odp"))
            {

                //Saving the PPTX presentation to PPTX format
                pres.Save("Aspose.pptx”,SaveFormat.Pptx", SaveFormat.Pptx);
            }

        }
    }
}
