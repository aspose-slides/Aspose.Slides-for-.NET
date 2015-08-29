using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Create a presentation
            Presentation pres = new Presentation();

            //Add the title slide
            Slide slide = pres.AddTitleSlide();

            //Set the title text
            ((TextHolder)slide.Placeholders[0]).Text = "Slide Title Heading";

            //Set the sub title text
            ((TextHolder)slide.Placeholders[1]).Text = "Slide Title Sub-Heading";

            //Write output to disk
            pres.Write("outAsposeSlides.ppt");
        }
    }
}
