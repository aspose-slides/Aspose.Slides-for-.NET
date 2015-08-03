using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace VSTO_Presentation
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            //Add a blank slide
            PowerPoint.Slide sld = Application.ActivePresentation.Slides[0];

            //Add a 15 x 15 table
            PowerPoint.Shape shp = sld.Shapes.AddTable(15, 15, 10, 10, 200, 300);
            PowerPoint.Table tbl = shp.Table;
            int i = -1;
            int j = -1;

            //Loop through all the rows
            foreach (PowerPoint.Row row in tbl.Rows)
            {
                i = i + 1;
                j = -1;

                //Loop through all the cells in the row
                foreach (PowerPoint.Cell cell in row.Cells)
                {
                    j = j + 1;
                    //Get text frame of each cell
                    PowerPoint.TextFrame tf = cell.Shape.TextFrame;
                    //Add some text
                    tf.TextRange.Text = "T" + i.ToString() + j.ToString();
                    //Set font size of the text as 10
                    tf.TextRange.Paragraphs(0, tf.TextRange.Text.Length).Font.Size = 10;
                }
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
