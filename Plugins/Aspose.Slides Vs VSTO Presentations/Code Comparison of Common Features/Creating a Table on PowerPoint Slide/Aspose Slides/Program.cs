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

            //Access first slide
            Slide sld = pres.GetSlideByPosition(1);

            //Add a table
            Aspose.Slides.Table tbl = sld.Shapes.AddTable(50, 50, pres.SlideSize.Width - 100, pres.SlideSize.Height - 100, 15, 15);

            //Loop through rows
            for (int i = 0; i < tbl.RowsNumber; i++)
                //Loop through cells
                for (int j = 0; j < tbl.ColumnsNumber; j++)
                {
                    //Get text frame of each cell
                    TextFrame tf = tbl.GetCell(j, i).TextFrame;
                    //Add some text
                    tf.Text = "T" + i.ToString() + j.ToString();
                    //Set font size of 10
                    tf.Paragraphs[0].Portions[0].FontHeight = 10;
                    tf.Paragraphs[0].HasBullet = false;
                }

            //Write the presentation to the disk
            pres.Write("tblSLD.ppt");
        }
    }
}
