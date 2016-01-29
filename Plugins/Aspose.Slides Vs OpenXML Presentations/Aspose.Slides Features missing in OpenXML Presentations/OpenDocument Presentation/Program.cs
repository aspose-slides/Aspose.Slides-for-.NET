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
            string Path = @"E:\Aspose\Aspose Vs OpenXML\Files\";
            //Instantiate a Presentation object that represents a presentation file
            using (Presentation pres = new Presentation(Path + "Source.odp"))
            {

                //Saving the PPTX presentation to PPTX format
                pres.Save(Path + "Aspose.pptx”,SaveFormat.Pptx", SaveFormat.Pptx);
            }

        }
    }
}
