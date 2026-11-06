using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following code example shows how to print all cached values of all chart series.
*/

namespace CSharp.Charts
{
    public class CachedChartValueExample
    {
        public static void Run()
        {
            // Source file name
            string sourcePath = Path.Combine(RunExamples.GetDataDir_Charts(), "Presentation_BubbleSizeRepresentation.pptx");

            using (Presentation pres = new Presentation(sourcePath))
            {
                var chart = (IChart)pres.Slides[0].Shapes[0];
                foreach (var series in chart.ChartData.Series)
                {
                    // Print all cached values of the chart series.
                    foreach (var point in series.DataPoints)
                    {
                        PrintCachedValue(point.CachedValue);
                        PrintCachedValue(point.CachedXValue);
                        PrintCachedValue(point.CachedYValue);
                        PrintCachedValue(point.CachedBubbleSize);
                    }
                }
            }
        }

        static void PrintCachedValue(ICachedChartValue cachedValue)
        {
            if (cachedValue.Data == null)
                Console.WriteLine("No cached data");

            double value = cachedValue.ToDouble();
            if (!double.IsNaN(value))
                Console.WriteLine("Numeric cached data: " + value);

            Console.WriteLine("String cached data: " + cachedValue.ToString());
        }
    }
}