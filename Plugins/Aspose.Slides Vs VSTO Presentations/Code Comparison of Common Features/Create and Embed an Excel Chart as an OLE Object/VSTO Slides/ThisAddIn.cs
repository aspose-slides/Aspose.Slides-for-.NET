using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using pptNS = Microsoft.Office.Interop.PowerPoint;
using xlNS = Microsoft.Office.Interop.Excel;

namespace VSTO_Slides
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            CreateNewChartInExcel();
            UseCopyPaste();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        public void SetCellValue(xlNS.Worksheet targetSheet, string Cell, object Value)
        {
            targetSheet.get_Range(Cell, Cell).set_Value(xlNS.XlRangeValueDataType.xlRangeValueDefault, Value);
        }

        public void CreateNewChartInExcel()
        {
            // Declare a variable for the Excel ApplicationClass instance.
            Microsoft.Office.Interop.Excel.Application excelApplication = new xlNS.Application();//new Microsoft.Office.Interop.Excel.ApplicationClass();

            // Declare variables for the Workbooks.Open method parameters. 
            string paramWorkbookPath = System.Windows.Forms.Application.StartupPath + @"\ChartData.xlsx";
            object paramMissing = Type.Missing;

            // Declare variables for the Chart.ChartWizard method.
            object paramChartFormat = 1;
            object paramCategoryLabels = 0;
            object paramSeriesLabels = 0;
            bool paramHasLegend = true;
            object paramTitle = "Sales by Quarter";
            object paramCategoryTitle = "Fiscal Quarter";
            object paramValueTitle = "Billions";

            try
            {
                // Create an instance of the Excel ApplicationClass object.          
                // excelApplication = new Microsoft.Office.Interop.Excel.ApplicationClass();

                // Create a new workbook with 1 sheet in it.
                xlNS.Workbook newWorkbook = excelApplication.Workbooks.Add(xlNS.XlWBATemplate.xlWBATWorksheet);

                // Change the name of the sheet.
                xlNS.Worksheet targetSheet = (xlNS.Worksheet)(newWorkbook.Worksheets[1]);
                targetSheet.Name = "Quarterly Sales";

                // Insert some data for the chart into the sheet.
                //              A       B       C       D       E
                //     1                Q1      Q2      Q3      Q4
                //     2    N. America  1.5     2       1.5     2.5
                //     3    S. America  2       1.75    2       2
                //     4    Europe      2.25    2       2.5     2
                //     5    Asia        2.5     2.5     2       2.75

                SetCellValue(targetSheet, "A2", "N. America");
                SetCellValue(targetSheet, "A3", "S. America");
                SetCellValue(targetSheet, "A4", "Europe");
                SetCellValue(targetSheet, "A5", "Asia");

                SetCellValue(targetSheet, "B1", "Q1");
                SetCellValue(targetSheet, "B2", 1.5);
                SetCellValue(targetSheet, "B3", 2);
                SetCellValue(targetSheet, "B4", 2.25);
                SetCellValue(targetSheet, "B5", 2.5);

                SetCellValue(targetSheet, "C1", "Q2");
                SetCellValue(targetSheet, "C2", 2);
                SetCellValue(targetSheet, "C3", 1.75);
                SetCellValue(targetSheet, "C4", 2);
                SetCellValue(targetSheet, "C5", 2.5);

                SetCellValue(targetSheet, "D1", "Q3");
                SetCellValue(targetSheet, "D2", 1.5);
                SetCellValue(targetSheet, "D3", 2);
                SetCellValue(targetSheet, "D4", 2.5);
                SetCellValue(targetSheet, "D5", 2);

                SetCellValue(targetSheet, "E1", "Q4");
                SetCellValue(targetSheet, "E2", 2.5);
                SetCellValue(targetSheet, "E3", 2);
                SetCellValue(targetSheet, "E4", 2);
                SetCellValue(targetSheet, "E5", 2.75);

                // Get the range holding the chart data.
                xlNS.Range dataRange = targetSheet.get_Range("A1", "E5");

                // Get the ChartObjects collection for the sheet.
                xlNS.ChartObjects chartObjects = (xlNS.ChartObjects)(targetSheet.ChartObjects(paramMissing));

                // Add a Chart to the collection.
                xlNS.ChartObject newChartObject = chartObjects.Add(0, 100, 600, 300);
                newChartObject.Name = "Sales Chart";

                // Create a new chart of the data.
                newChartObject.Chart.ChartWizard(dataRange, xlNS.XlChartType.xl3DColumn, paramChartFormat, xlNS.XlRowCol.xlRows,
                    paramCategoryLabels, paramSeriesLabels, paramHasLegend, paramTitle, paramCategoryTitle, paramValueTitle, paramMissing);

                // Save the workbook.
                newWorkbook.SaveAs(paramWorkbookPath, paramMissing, paramMissing, paramMissing, paramMissing,
                    paramMissing, xlNS.XlSaveAsAccessMode.xlNoChange, paramMissing, paramMissing, paramMissing, paramMissing, paramMissing);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (excelApplication != null)
                {
                    // Close Excel.
                    excelApplication.Quit();
                }
            }
        }

        public void UseCopyPaste()
        {
            // Declare variables to hold references to PowerPoint objects.
            pptNS.Application powerpointApplication = null;
            pptNS.Presentation pptPresentation = null;
            pptNS.Slide pptSlide = null;
            pptNS.ShapeRange shapeRange = null;

            // Declare variables to hold references to Excel objects.
            xlNS.Application excelApplication = null;
            xlNS.Workbook excelWorkBook = null;
            xlNS.Worksheet targetSheet = null;
            xlNS.ChartObjects chartObjects = null;
            xlNS.ChartObject existingChartObject = null;

            string paramPresentationPath = System.Windows.Forms.Application.StartupPath + @"\ChartTest.pptx";
            string paramWorkbookPath = System.Windows.Forms.Application.StartupPath + @"\ChartData.xlsx";
            object paramMissing = Type.Missing;

            try
            {
                // Create an instance of PowerPoint.
                powerpointApplication = new pptNS.Application();

                // Create an instance Excel.          
                excelApplication = new xlNS.Application();

                // Open the Excel workbook containing the worksheet with the chart data.
                excelWorkBook = excelApplication.Workbooks.Open(paramWorkbookPath,
                    paramMissing, paramMissing, paramMissing, paramMissing, paramMissing,
                    paramMissing, paramMissing, paramMissing, paramMissing, paramMissing,
                    paramMissing, paramMissing, paramMissing, paramMissing);

                // Get the worksheet that contains the chart.
                targetSheet =
                    (xlNS.Worksheet)(excelWorkBook.Worksheets["Quarterly Sales"]);

                // Get the ChartObjects collection for the sheet.
                chartObjects =
                    (xlNS.ChartObjects)(targetSheet.ChartObjects(paramMissing));

                // Get the chart to copy.
                existingChartObject =
                    (xlNS.ChartObject)(chartObjects.Item("Sales Chart"));

                // Create a PowerPoint presentation.
                pptPresentation =
                    powerpointApplication.Presentations.Add(
                    Microsoft.Office.Core.MsoTriState.msoTrue);

                // Add a blank slide to the presentation.
                pptSlide =
                    pptPresentation.Slides.Add(1, pptNS.PpSlideLayout.ppLayoutBlank);

                // Copy the chart from the Excel worksheet to the clipboard.
                existingChartObject.Copy();

                // Paste the chart into the PowerPoint presentation.
                shapeRange = pptSlide.Shapes.Paste();

                // Position the chart on the slide.
                shapeRange.Left = 60;
                shapeRange.Top = 100;

                // Save the presentation.
                pptPresentation.SaveAs(paramPresentationPath, pptNS.PpSaveAsFileType.ppSaveAsOpenXMLPresentation, Microsoft.Office.Core.MsoTriState.msoTrue);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Release the PowerPoint slide object.
                shapeRange = null;
                pptSlide = null;

                // Close and release the Presentation object.
                if (pptPresentation != null)
                {
                    pptPresentation.Close();
                    pptPresentation = null;
                }

                // Quit PowerPoint and release the ApplicationClass object.
                if (powerpointApplication != null)
                {
                    powerpointApplication.Quit();
                    powerpointApplication = null;
                }

                // Release the Excel objects.
                targetSheet = null;
                chartObjects = null;
                existingChartObject = null;

                // Close and release the Excel Workbook object.
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }

                // Quit Excel and release the ApplicationClass object.
                if (excelApplication != null)
                {
                    excelApplication.Quit();
                    excelApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
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
