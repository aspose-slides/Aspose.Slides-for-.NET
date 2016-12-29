using System.IO;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class NumberFormat
    {
        public static void Run()
        {
            //ExStart:NumberFormat
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate the presentation// Instantiate the presentation
            Presentation pres = new Presentation();

            // Access the first presentation slide
            ISlide slide = pres.Slides[0];

            // Adding a defautlt clustered column chart
            IChart chart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400);

            // Accessing the chart series collection
            IChartSeriesCollection series = chart.ChartData.Series;

            // Setting the preset number format
            // Traverse through every chart series
            foreach (ChartSeries ser in series)
            {
                // Traverse through every data cell in series
                foreach (IChartDataPoint cell in ser.DataPoints)
                {
                    // Setting the number format
                    cell.Value.AsCell.PresetNumberFormat = 10; //0.00%
                }
            }

            // Saving presentation
            pres.Save(dataDir + "PresetNumberFormat_out.pptx", SaveFormat.Pptx);
            //ExEnd:NumberFormat
        }
    }
}