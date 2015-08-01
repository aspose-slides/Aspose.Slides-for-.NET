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
            //Instantiate Prseetation class that represents the PPTX
            Presentation pres = new Presentation();

            //Get the first slide
            ISlide slide = pres.Slides[0];

            //Add an autoshape of type line
            slide.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0);
        }
    }
}
