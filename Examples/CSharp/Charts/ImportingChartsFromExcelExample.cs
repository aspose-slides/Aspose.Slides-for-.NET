using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides.Excel;
using Aspose.Slides.Import;

/*
The following example demonstrates how to import all charts from an Excel workbook to a presentation. 
*/

namespace CSharp.Charts
{
    class ImportingChartsFromExcelExample
    {
        public static void Run()
        {
            // Path to excel file
            string dataDir = RunExamples.GetDataDir_Charts();
            string externalWbPath = dataDir + "book1.xlsx";

            // Path to output file
            string outFileName = Path.Combine(RunExamples.OutPath, "ImportExcelChart.pptx");

            // Initializes a new instance using the specified file path
            ExcelDataWorkbook workbook = new ExcelDataWorkbook(externalWbPath);

            using (var pres = new Presentation())
            {
                var blankLayout = pres.LayoutSlides.GetByType(SlideLayoutType.Blank);

                // Gets the names of all worksheets contained in the Excel workbook
                var worksheetNames = workbook.GetWorksheetNames();
                foreach (var name in worksheetNames)
                {
                    // Gets a dictionary containing the indexes and names of all charts in the specified worksheet of an Excel workbook
                    var worksheetCharts = workbook.GetChartsFromWorksheet(name);
                    foreach (var chart in worksheetCharts)
                    {
                        ISlide slide = pres.Slides.AddEmptySlide(blankLayout);
                        // Imports the chart from a workbook file by its name and worksheet name
                        ExcelWorkbookImporter.AddChartFromWorkbook(slide.Shapes, 10, 10, workbook, name, chart.Key, false);
                    }
                }

                // Saves result
                pres.Save(outFileName, SaveFormat.Pptx);
            }
        }
    }
}
