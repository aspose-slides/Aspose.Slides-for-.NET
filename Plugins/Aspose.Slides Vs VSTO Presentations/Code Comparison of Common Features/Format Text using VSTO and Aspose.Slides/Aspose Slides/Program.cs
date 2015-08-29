using Aspose.Slides;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Aspose_Slides
{
    class Program
    {
        static void Main(string[] args)
        {
            //Open the presentation
            Presentation pres = new Presentation("source.ppt");

            //Add Verdana font
            FontEntity font = pres.Fonts[0];
            FontEntity verdanaFont = new FontEntity(pres, font);
            verdanaFont.FontName = "Verdana";
            int verdanaFontIndex = pres.Fonts.Add(verdanaFont);

            //Access the first slide
            Slide slide = pres.GetSlideByPosition(1);

            //Access the third shape
            Shape shp = slide.Shapes[2];

            //Change its text's font to Verdana and height to 32
            TextFrame tf = shp.TextFrame;
            Paragraph para = tf.Paragraphs[0];
            Portion port = para.Portions[0];
            port.FontIndex = verdanaFontIndex;
            port.FontHeight = 32;

            //Bolden it
            port.FontBold = true;

            //Italicize it
            port.FontItalic = true;

            //Change text color
            port.FontColor = Color.FromArgb(0x33, 0x33, 0xCC);

            //Change shape background color
            shp.FillFormat.Type = FillType.Solid;
            shp.FillFormat.ForeColor = Color.FromArgb(0xCC, 0xCC, 0xFF);

            //Write the output to disk
            pres.Write("outAspose.ppt");
        }
    }
}
