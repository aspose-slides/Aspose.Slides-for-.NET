using Aspose.Slides;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CSharp.Tables
{
    class CreateTable
    {
        public static void Run() {

            //ExStart:CreateTable

            Presentation pres = new Presentation();

            //Access first slide
            ISlide sld = pres.Slides[0];

            //Define columns with widths and rows with heights
            double[] dblCols = { 50, 50, 50 };
            double[] dblRows = { 50, 30, 30, 30, 30 };

            //Add a table
            Aspose.Slides.ITable tbl = sld.Shapes.AddTable(50, 50, dblCols, dblRows);

            //Set border format for each cell
            foreach (IRow row in tbl.Rows)
            {
                foreach (ICell cell in row)
                {

                    //Get text frame of each cell
                    ITextFrame tf = cell.TextFrame;
                    //Add some text
                    tf.Text = "T" + cell.FirstRowIndex.ToString() + cell.FirstColumnIndex.ToString();
                    //Set font size of 10
                    tf.Paragraphs[0].Portions[0].PortionFormat.FontHeight = 10;
                    tf.Paragraphs[0].ParagraphFormat.Bullet.Type = BulletType.None;
                }
            }

            //Write the presentation to the disk
            pres.Save("C:\\data\\tblSLD.ppt", SaveFormat.Ppt);
            //ExEnd:CreateTable

        }
    }
}
