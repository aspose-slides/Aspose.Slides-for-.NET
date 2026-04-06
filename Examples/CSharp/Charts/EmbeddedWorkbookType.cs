using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following code example shows how to check the format of a chart's workbook before working with the chart's data.
*/

namespace CSharp.Charts
{
    public class EmbeddedWorkbookType
    {
        public static void Run()
        {
            // Source file name
            string sourcePath = Path.Combine(RunExamples.GetDataDir_Charts(), "EmbeddedWorkbook.pptx");
            // Output file name
            string resultPath = Path.Combine(RunExamples.OutPath, "EmbeddedWorkbook-out.pptx");

            using (var presentation = new Presentation(sourcePath))
            {
                foreach (var shape in presentation.Slides[0].Shapes)
                {
                    if (!(shape is IChart chart))
                        continue;

                    var chartData = chart.ChartData;

                    // Skip charts whose embedded workbook format is not supported.
                    if (chartData.DataSourceType == ChartDataSourceType.InternalWorkbook &&
                        chartData.EmbeddedWorkbookType == WorkbookType.WorkbookBinaryMacro)
                    {
                        Console.WriteLine("\nSkip charts whose embedded workbook format is {0}", chartData.EmbeddedWorkbookType);
                        continue;
                    }

                    Console.WriteLine("\nWork with charts whose embedded workbook format is {0}:", chartData.EmbeddedWorkbookType);

                    // Read or modify chart workbook data.
                    Console.WriteLine("\tChart old data: {0}", chartData.Series[0].Name.AsCells.GetHashCode());

                    var cell = chartData.Series[0].DataPoints[0].Value.AsCell;
                    Console.WriteLine("\tChart new data: {0}", cell.Value);
                }

                presentation.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}