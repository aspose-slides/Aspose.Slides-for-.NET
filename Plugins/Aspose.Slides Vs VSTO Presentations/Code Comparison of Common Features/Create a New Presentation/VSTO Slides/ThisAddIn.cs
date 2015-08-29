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

            PowerPoint.Presentation pres = Globals.ThisAddIn.Application
                .Presentations.Add(Microsoft.Office.Core.MsoTriState.msoFalse);

            //Get the title slide layout
            PowerPoint.CustomLayout layout = pres.SlideMaster.
                CustomLayouts[PowerPoint.PpSlideLayout.ppLayoutTitle];

            //Add a title slide.
            PowerPoint.Slide slide = pres.Slides.AddSlide(1, layout);

            //Set the title text
            slide.Shapes.Title.TextFrame.TextRange.Text = "Slide Title Heading";

            //Set the sub title text
            slide.Shapes[2].TextFrame.TextRange.Text = "Slide Title Sub-Heading";

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
