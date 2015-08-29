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
            findReplaceText("Aspose", "Aspose for .Net");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        private void findReplaceText(string strToFind, string strToReplaceWith)
        {

            //Open the presentation
            PowerPoint.Presentation pres = null;
            pres = Globals.ThisAddIn.Application.Presentations.Open("mytextone.ppt",
                                      Microsoft.Office.Core.MsoTriState.msoFalse,
                                      Microsoft.Office.Core.MsoTriState.msoFalse,
                                      Microsoft.Office.Core.MsoTriState.msoFalse);

            //Loop through slides
            foreach (PowerPoint.Slide sld in pres.Slides)
                //Loop through all shapes in slide
                foreach (PowerPoint.Shape shp in sld.Shapes)
                {
                    //Access text in the shape
                    string str = shp.TextFrame.TextRange.Text;
                    //Find text to replace
                    if (str.Contains(strToFind))
                    //Replace exisitng text with the new text
                    {
                        int idx = str.IndexOf(strToFind);
                        string strStartText = str.Substring(0, idx);
                        string strEndText = str.Substring(idx + strToFind.Length, str.Length - 1 - (idx + strToFind.Length - 1));
                        shp.TextFrame.TextRange.Text = strStartText + strToReplaceWith + strEndText;
                    }
                    pres.SaveAs("MyTextOne___.ppt",
                    PowerPoint.PpSaveAsFileType.ppSaveAsPresentation,
                    Microsoft.Office.Core.MsoTriState.msoFalse);
                }
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
