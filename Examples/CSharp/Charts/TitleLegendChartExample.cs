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
The following example demonstrates how to get an actual layout of a chart. 
*/

namespace CSharp.Charts
{
    class TitleLegendChartExample
    {
        public static void Run()
        {
            using (Presentation pres = new Presentation())
            {
                var chart = (Chart)pres.Slides[0].Shapes.AddChart(ChartType.ClusteredColumn, 100, 100, 500, 350);
                chart.ValidateChartLayout();

                var chartTitle = chart.ChartTitle;
                Console.WriteLine($"ChartTitle.X = {chartTitle.ActualX}, ChartTitle.Y = {chartTitle.ActualY}");
                Console.WriteLine($"ChartTitle.Width = {chartTitle.ActualWidth}, ChartTitle.Height = {chartTitle.ActualHeight}");

                var legend = chart.Legend;
                Console.WriteLine($"Legend.X = {legend.ActualX}, Legend.Y = {legend.ActualY}");
                Console.WriteLine($"Legend.Width = {legend.ActualWidth}, Legend.Height = {legend.ActualHeight}");
            }
        }
    }
}
