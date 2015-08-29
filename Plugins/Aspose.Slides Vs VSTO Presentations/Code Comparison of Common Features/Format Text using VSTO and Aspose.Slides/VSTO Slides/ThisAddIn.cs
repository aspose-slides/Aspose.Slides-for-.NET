using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;

namespace VSTO_Slides
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            PowerPoint.Presentation pres = null;

            //Open the presentation
            pres = Globals.ThisAddIn.Application.Presentations.Open("source.ppt",
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoFalse,
                Microsoft.Office.Core.MsoTriState.msoTrue);

            //Access the first slide
            PowerPoint.Slide slide = pres.Slides[1];

            //Access the third shape
            PowerPoint.Shape shp = slide.Shapes[3];

            //Change its text's font to Verdana and height to 32
            PowerPoint.TextRange txtRange = shp.TextFrame.TextRange;
            txtRange.Font.Name = "Verdana";
            txtRange.Font.Size = 32;

            //Bolden it
            txtRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoCTrue;

            //Italicize it
            txtRange.Font.Italic = Microsoft.Office.Core.MsoTriState.msoCTrue;

            //Change text color
            txtRange.Font.Color.RGB = 0x00CC3333;

            //Change shape background color
            shp.Fill.ForeColor.RGB = 0x00FFCCCC;

            //Reposition it horizontally
            shp.Left -= 70;

            //Write the output to disk
            pres.SaveAs("outVSTO.ppt",
                PowerPoint.PpSaveAsFileType.ppSaveAsPresentation,
                Microsoft.Office.Core.MsoTriState.msoFalse);
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
