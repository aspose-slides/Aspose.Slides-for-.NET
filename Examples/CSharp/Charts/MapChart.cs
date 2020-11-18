using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
Map charts example.
*/
namespace CSharp.Charts
{
    public class MapChart
    {
        // This example demonstrates creating Map charts.
        // Please pay attension that when you first open a presentation in PP it may take a few seconds to upload an image 
        // of the chart from the Bing service since we don't provide cached image.

        public static void Run()
        {
            string resultPath = Path.Combine(RunExamples.OutPath, "MapChart_out.pptx");

            using (Presentation presentation = new Presentation())
            {
                //create empty chart
                IChart chart = presentation.Slides[0].Shapes.AddChart(ChartType.Map, 50, 50, 500, 400, false);

                IChartDataWorkbook wb = chart.ChartData.ChartDataWorkbook;

                //Add series and few data points
                IChartSeries series = chart.ChartData.Series.Add(ChartType.Map);
                series.DataPoints.AddDataPointForMapSeries(wb.GetCell(0, "B2", 5));
                series.DataPoints.AddDataPointForMapSeries(wb.GetCell(0, "B3", 1));
                series.DataPoints.AddDataPointForMapSeries(wb.GetCell(0, "B4", 10));

                //add categories
                chart.ChartData.Categories.Add(wb.GetCell(0, "A2", "United States"));
                chart.ChartData.Categories.Add(wb.GetCell(0, "A3", "Mexico"));
                chart.ChartData.Categories.Add(wb.GetCell(0, "A4", "Brazil"));

                //change data point value    
                IChartDataPoint dataPoint = series.DataPoints[1];
                dataPoint.ColorValue.AsCell.Value = "15";

                //set data point appearance    
                dataPoint.Format.Fill.FillType = FillType.Solid;
                dataPoint.Format.Fill.SolidFillColor.Color = Color.Green;

                presentation.Save(resultPath, SaveFormat.Pptx);
            }
        }
    }
}
