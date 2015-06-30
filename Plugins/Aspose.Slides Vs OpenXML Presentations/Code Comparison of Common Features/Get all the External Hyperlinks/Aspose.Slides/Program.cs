using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose.Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            string fileName = @"E:\Aspose\Aspose Vs OpenXML\Aspose.Slides Vs OpenXML Presentation v1.1\Sample Files\Get all the External Eyperlinks.pptx";
            //Instantiate a Presentation object that represents a PPTX file
            Presentation pres = new Presentation(fileName);

            //Get the hyperlinks from presentation
            IList<IHyperlinkContainer> links = pres.HyperlinkQueries.GetAnyHyperlinks();

            foreach (IHyperlinkContainer link in links)
                Console.WriteLine(link.HyperlinkClick.ExternalUrl);

        }
    }
}
