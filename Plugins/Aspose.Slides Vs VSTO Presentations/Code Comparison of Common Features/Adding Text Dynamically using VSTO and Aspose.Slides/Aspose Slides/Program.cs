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

            //Blank slide is added by default, when you create
            //presentation from default constructor
            //So, we don't need to add any blank slide
            Slide sld = pres.GetSlideByPosition(1);

            //Get the font index for Arial 
            //It is always 0 if you create presentation from
            //default constructor
            int arialFontIndex = 0;

            //Add a textbox
            //To add it, we will first add a rectangle
            Shape shp = sld.Shapes.AddRectangle(1200, 800, 3200, 370);

            //Hide its line
            shp.LineFormat.ShowLines = false;

            //Then add a textframe inside it
            TextFrame tf = shp.AddTextFrame("");

            //Set a text
            tf.Text = "Text added dynamically";
            Portion port = tf.Paragraphs[0].Portions[0];
            port.FontIndex = arialFontIndex;
            port.FontBold = true;
            port.FontHeight = 32;

            //Write the output to disk
            pres.Write("outAspose.ppt");
        }
    }
}
