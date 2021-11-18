using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
This example demonstrates how to use the spreadsheet options for a chart formulas.
*/
namespace CSharp.Charts
{
    class SpreadsheetFormulasOptions
    {
        public static void Run()
        {
            LoadOptions loadOptions = new LoadOptions();

            // Set preferred culture information for calculating some functions intended for use with languages 
            // that use the double-byte character set (DBCS).
            loadOptions.SpreadsheetOptions.PreferredCulture = new System.Globalization.CultureInfo("ja-JP");

            using (Presentation presentation = new Presentation(loadOptions))
            {
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 150, 150, 500, 300);
                IChartDataWorkbook workbook = chart.ChartData.ChartDataWorkbook;

                var cell = workbook.GetCell(0, "B2");
                
                // Use the Formula property of the IChartDataCell interface to write a formula in a cell.
                cell.Formula = "FINDB(\"ス\", \"テキスト\")";
                workbook.CalculateFormulas();

                //Check calculation.
                if (Int32.Parse(cell.Value.ToString()) == 5)
                {
                    Console.WriteLine("Calculated value = 5.");
                }
                else
                {
                    Console.WriteLine("Wrong calculation!");
                }
            }
        }
    }
}
