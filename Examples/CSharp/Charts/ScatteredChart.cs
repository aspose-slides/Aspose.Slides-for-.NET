using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class ScatteredChart
    {
        public static void Run()
        {
            //ExStart:ScatteredChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            Presentation pres = new Presentation();

            ISlide slide = pres.Slides[0];

            // Creating the default chart
            IChart chart = slide.Shapes.AddChart(ChartType.ScatterWithSmoothLines, 0, 0, 400, 400);

            // Getting the default chart data worksheet index
            int defaultWorksheetIndex = 0;

            // Getting the chart data worksheet
            IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

            // Delete demo series
            chart.ChartData.Series.Clear();

            // Add new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, "Series 1"), chart.Type);
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 3, "Series 2"), chart.Type);

            // Take first chart series
            IChartSeries series = chart.ChartData.Series[0];

            // Add new point (1:3) there.
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 1), fact.GetCell(defaultWorksheetIndex, 2, 2, 3));

            // Add new point (2:10)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 2), fact.GetCell(defaultWorksheetIndex, 3, 2, 10));

            // Edit the type of series
            series.Type = ChartType.ScatterWithStraightLinesAndMarkers;

            // Changing the chart series marker
            series.Marker.Size = 10;
            series.Marker.Symbol = MarkerStyleType.Star;

            // Take second chart series
            series = chart.ChartData.Series[1];

            // Add new point (5:2) there.
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 2, 3, 5), fact.GetCell(defaultWorksheetIndex, 2, 4, 2));

            // Add new point (3:1)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 3, 3, 3), fact.GetCell(defaultWorksheetIndex, 3, 4, 1));

            // Add new point (2:2)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 4, 3, 2), fact.GetCell(defaultWorksheetIndex, 4, 4, 2));

            // Add new point (5:1)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 5, 3, 5), fact.GetCell(defaultWorksheetIndex, 5, 4, 1));

            // Changing the chart series marker
            series.Marker.Size = 10;
            series.Marker.Symbol = MarkerStyleType.Circle;

            pres.Save(dataDir + "AsposeChart_out.pptx", SaveFormat.Pptx);
            //ExEnd:ScatteredChart
        }
    }
}