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
            //Create a presentation
            PowerPoint.Presentation pres = Globals.ThisAddIn.Application
                .Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);

            //Get the blank slide layout
            PowerPoint.CustomLayout layout = pres.SlideMaster.
                CustomLayouts[7];

            //Add a blank slide
            PowerPoint.Slide sld = pres.Slides.AddSlide(1, layout);

            //Add a text
            PowerPoint.Shape shp = sld.Shapes.AddTextbox(Microsoft.Office.Core.MsoTextOrientation.msoTextOrientationHorizontal, 150, 100, 400, 100);

            //Set a text
            PowerPoint.TextRange txtRange = shp.TextFrame.TextRange;
            txtRange.Text = "Text added dynamically";
            txtRange.Font.Name = "Arial";
            txtRange.Font.Bold = Microsoft.Office.Core.MsoTriState.msoTrue;
            txtRange.Font.Size = 32;

            //Write the output to disk
            pres.SaveAs("outVSTOAddingText.ppt",
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
