using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Excel;
using Aspose.Slides.Export;
using Aspose.Slides.Import;

/*
This example demonstrates how to import tables from an Excel workbook
by specifying a worksheet name and cell range.
*/

namespace Aspose.Slides.Examples.CSharp.Tables
{
    public class AddTableFromWorkbookExample
    {
        public static void Run()
        {
            // Path to the source Excel file
            string excelFilePath = Path.Combine(RunExamples.GetDataDir_Tables(), "Budget.xlsx");
            // Path to the output presentation
            string outPath = Path.Combine(RunExamples.OutPath, "TableFromWorkbook.pptx");

            using (var presentation = new Presentation())
            {
                // Get the layout of the first slide to reuse when adding new slides
                ILayoutSlide slideLayout = presentation.Slides[0].LayoutSlide;

                // Create workbook instance from file
                IExcelDataWorkbook workbook = new ExcelDataWorkbook(excelFilePath);
                // Import the table using an IExcelDataWorkbook instance
                ExcelWorkbookImporter.AddTableFromWorkbook(presentation.Slides[0].Shapes, 10, 10, workbook, "Month", "D4:H17");

                // Add a new slide
                ISlide secondSlide = presentation.Slides.AddEmptySlide(slideLayout);
                // Import the table directly from an Excel file path
                ExcelWorkbookImporter.AddTableFromWorkbook(secondSlide.Shapes, 10, 10, excelFilePath, "Budget", "B21:E43");

                // Add a new slide
                ISlide thirdSlide = presentation.Slides.AddEmptySlide(slideLayout);
                // Import the table from an Excel stream
                using (FileStream fStream = new FileStream(excelFilePath, FileMode.Open, FileAccess.Read))
                {
                    ExcelWorkbookImporter.AddTableFromWorkbook(thirdSlide.Shapes, 10, 10, fStream,
                        "Budget", "B47:E55");
                }

                // Save the presentation
                presentation.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}