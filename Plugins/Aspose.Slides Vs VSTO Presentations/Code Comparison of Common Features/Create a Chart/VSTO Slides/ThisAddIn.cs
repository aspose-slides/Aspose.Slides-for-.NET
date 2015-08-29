using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Drawing;
using Microsoft.Office.Core;

namespace VSTO_Slides
{
    public partial class ThisAddIn
    {
        //Global Variables
        public static Microsoft.Office.Interop.PowerPoint.Application objPPT;
        public static Microsoft.Office.Interop.PowerPoint.Presentation objPres;

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {

            EnsurePowerPointIsRunning(true, true);

            //Instantiate slide object
            Microsoft.Office.Interop.PowerPoint.Slide objSlide = null;

            //Access the first slide of presentation
            objSlide = objPres.Slides[1];

            //Select firs slide and set its layout
            objSlide.Select();
            objSlide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutBlank;

            //Add a default chart in slide
            objSlide.Shapes.AddChart(Microsoft.Office.Core.XlChartType.xl3DColumn, 20F, 30F, 400F, 300F);

            //Access the added chart
            Microsoft.Office.Interop.PowerPoint.Chart ppChart = objSlide.Shapes[1].Chart;

            //Access the chart data
            Microsoft.Office.Interop.PowerPoint.ChartData chartData = ppChart.ChartData;

            //Create instance to Excel workbook to work with chart data
            Microsoft.Office.Interop.Excel.Workbook dataWorkbook = (Microsoft.Office.Interop.Excel.Workbook)chartData.Workbook;

            //Accessing the data worksheet for chart
            Microsoft.Office.Interop.Excel.Worksheet dataSheet = dataWorkbook.Worksheets[1];

            //Setting the range of chart
            Microsoft.Office.Interop.Excel.Range tRange = dataSheet.Cells.get_Range("A1", "B5");

            //Applying the set range on chart data table
            Microsoft.Office.Interop.Excel.ListObject tbl1 = dataSheet.ListObjects["Table1"];
            tbl1.Resize(tRange);

            //Setting values for categories and respective series data

            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A2"))).FormulaR1C1 = "Bikes";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A3"))).FormulaR1C1 = "Accessories";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A4"))).FormulaR1C1 = "Repairs";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("A5"))).FormulaR1C1 = "Clothing";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B2"))).FormulaR1C1 = "1000";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B3"))).FormulaR1C1 = "2500";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B4"))).FormulaR1C1 = "4000";
            ((Microsoft.Office.Interop.Excel.Range)(dataSheet.Cells.get_Range("B5"))).FormulaR1C1 = "3000";

            //Setting chart title
            ppChart.ChartTitle.Font.Italic = true;
            ppChart.ChartTitle.Text = "2007 Sales";
            ppChart.ChartTitle.Font.Size = 18;
            ppChart.ChartTitle.Font.Color = Color.Black.ToArgb();
            ppChart.ChartTitle.Format.Line.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            ppChart.ChartTitle.Format.Line.ForeColor.RGB = Color.Black.ToArgb();

            //Accessing Chart value axis
            Microsoft.Office.Interop.PowerPoint.Axis valaxis = ppChart.Axes(Microsoft.Office.Interop.PowerPoint.XlAxisType.xlValue, Microsoft.Office.Interop.PowerPoint.XlAxisGroup.xlPrimary);

            //Setting values axis units
            valaxis.MajorUnit = 2000.0F;
            valaxis.MinorUnit = 1000.0F;
            valaxis.MinimumScale = 0.0F;
            valaxis.MaximumScale = 4000.0F;

            //Accessing Chart Depth axis
            Microsoft.Office.Interop.PowerPoint.Axis Depthaxis = ppChart.Axes(Microsoft.Office.Interop.PowerPoint.XlAxisType.xlSeriesAxis, Microsoft.Office.Interop.PowerPoint.XlAxisGroup.xlPrimary);
            Depthaxis.Delete();

            //Setting chart rotation
            ppChart.Rotation = 20; //Y-Value
            ppChart.Elevation = 15; //X-Value
            ppChart.RightAngleAxes = false;

            // Save the presentation as a PPTX
            objPres.SaveAs("VSTOSampleChart", Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsDefault, MsoTriState.msoTrue);

            //Close Workbook and presentation
            dataWorkbook.Application.Quit();
            objPres.Application.Quit();

        }

        //Supplementary methods
        public static void StartPowerPoint()
        {
            objPPT = new Microsoft.Office.Interop.PowerPoint.Application();
            objPPT.Visible = MsoTriState.msoTrue;
            //  objPPT.WindowState = PowerPoint.PpWindowState.ppWindowMaximized
        }

        public static void EnsurePowerPointIsRunning(bool blnAddPresentation)
        {
            EnsurePowerPointIsRunning(blnAddPresentation, false);
        }

        public static void EnsurePowerPointIsRunning()
        {
            EnsurePowerPointIsRunning(false, false);
        }

        public static void EnsurePowerPointIsRunning(bool blnAddPresentation, bool blnAddSlide)
        {
            string strName = null;
            //
            //Try accessing the name property. If it causes an exception then 
            //start a new instance of PowerPoint
            try
            {
                strName = objPPT.Name;
            }
            catch (Exception)
            {

                StartPowerPoint();
            }
            //
            //blnAddPresentation is used to ensure there is a presentation loaded
            if (blnAddPresentation == true)
            {
                try
                {
                    strName = objPres.Name;
                }
                catch (Exception)
                {
                    objPres = objPPT.Presentations.Add(MsoTriState.msoTrue);
                }
            }
            //
            //BlnAddSlide is used to ensure there is at least one slide in the 
            //presentation
            if (blnAddSlide)
            {
                try
                {
                    strName = objPres.Slides[1].Name;
                }
                catch (Exception)
                {
                    Microsoft.Office.Interop.PowerPoint.Slide objSlide = null;
                    Microsoft.Office.Interop.PowerPoint.CustomLayout objCustomLayout = null;
                    objCustomLayout = objPres.SlideMaster.CustomLayouts[1];
                    objSlide = objPres.Slides.AddSlide(1, objCustomLayout);
                    objSlide.Layout = Microsoft.Office.Interop.PowerPoint.PpSlideLayout.ppLayoutText;
                    objCustomLayout = null;
                    objSlide = null;
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
