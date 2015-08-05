using Microsoft.Office.Interop.PowerPoint;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace VSTO_Presentation
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //Open Prsentation class that contains the table
            string FileName = @"E:\Aspose\Aspose Vs VSTO\Aspose.Slides Vs VSTO Presentations v 1.1\Sample Files\Removing Row Or Column in Table.pptx";
            Presentation pres = Application.Presentations.Open(FileName);

            //Get the first slide
            Slide sld = pres.Slides[1];

            foreach (Shape shp in sld.Shapes)
            {
                if (shp.HasTable == Microsoft.Office.Core.MsoTriState.msoTrue)
                {
                    shp.Table.Rows[1].Delete();
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
