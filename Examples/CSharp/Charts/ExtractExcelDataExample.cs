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

/*
The following example demonstrates how to extract a value from a cell and how to retrieve 
worksheet names and chart names from an Excel workbook.
*/

namespace CSharp.Charts
{
    class ExtractExcelDataExample
    {
        public static void Run()
        {
            string dataDir = RunExamples.GetDataDir_Charts();
            string externalWbPath = dataDir + "book1.xlsx";

            // extract a value from a cell
            IExcelDataWorkbook workbook = new ExcelDataWorkbook(externalWbPath);
            IExcelDataCell cell = workbook.GetCell("Sheet2", "B2");
            Console.WriteLine(cell.Value);

            //retrieve worksheet names and chart names from an Excel workbook
            var sheetNames = workbook.GetWorksheetNames();
            foreach (var name in sheetNames)
            {
                Console.WriteLine("Worksheet " + name + " has the following charts:");

                var sheetCharts = workbook.GetChartsFromWorksheet(name);
                foreach (var chart in sheetCharts)
                    Console.WriteLine(chart.Key + " - " + chart.Value);
            }
        }
    }
}
