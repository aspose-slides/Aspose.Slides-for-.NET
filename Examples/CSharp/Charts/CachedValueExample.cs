using System;
using System.Drawing;
using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Examples.CSharp;
using Aspose.Slides.Export;

/*
The following code example shows how to print all cached chart category values.
 */

namespace CSharp.Charts
{
    public class CachedValueExample
    {
        public static void Run()
        {
            // Source file name
            string sourcePath = Path.Combine(RunExamples.GetDataDir_Charts(), "Presentation.pptx");

            using (Presentation pres = new Presentation(sourcePath))
            {
                var chart = (IChart)pres.Slides[0].Shapes[0];

                // Print all cached chart category values.
                foreach (var category in chart.ChartData.Categories)
                    Console.WriteLine(category.CachedValue.Data == null
                        ? "No cached data"
                        : category.CachedValue.ToString());
            }
        }
    }
}