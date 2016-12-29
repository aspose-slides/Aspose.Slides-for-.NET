using System.IO;

using Aspose.Slides;
using Aspose.Slides.Charts;
using System.Drawing;
using Aspose.Slides.Export;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class NormalCharts
    {
        public static void Run()
        {
            //ExStart:NormalCharts
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create directory if it is not already present.
            bool IsExists = System.IO.Directory.Exists(dataDir);
            if (!IsExists)
                System.IO.Directory.CreateDirectory(dataDir);

            // Instantiate Presentation class that represents PPTX file
            Presentation pres = new Presentation();

            // Access first slide
            ISlide sld = pres.Slides[0];

            // Add chart with default data
            IChart chart = sld.Shapes.AddChart(ChartType.ClusteredColumn, 0, 0, 500, 500);

            // Setting chart Title
            // Chart.ChartTitle.TextFrameForOverriding.Text = "Sample Title";
            chart.ChartTitle.AddTextFrameForOverriding("Sample Title");
            chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.True;
            chart.ChartTitle.Height = 20;
            chart.HasTitle = true;

            // Set first series to Show Values
            chart.ChartData.Series[0].Labels.DefaultDataLabelFormat.ShowValue = true;

            // Setting the index of chart data sheet
            int defaultWorksheetIndex = 0;

            // Getting the chart data worksheet
            IChartDataWorkbook fact = chart.ChartData.ChartDataWorkbook;

            // Delete default generated series and categories
            chart.ChartData.Series.Clear();
            chart.ChartData.Categories.Clear();
            int s = chart.ChartData.Series.Count;
            s = chart.ChartData.Categories.Count;

            // Adding new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.Type);
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.Type);

            // Adding new categories
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"));
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"));
            chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"));

            // Take first chart series
            IChartSeries series = chart.ChartData.Series[0];

            // Now populating series data

            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 20));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 50));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 30));

            // Setting fill color for series
            series.Format.Fill.FillType = FillType.Solid;
            series.Format.Fill.SolidFillColor.Color = Color.Red;


            // Take second chart series
            series = chart.ChartData.Series[1];

            // Now populating series data
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 2, 30));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 2, 10));
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 2, 60));

            // Setting fill color for series
            series.Format.Fill.FillType = FillType.Solid;
            series.Format.Fill.SolidFillColor.Color = Color.Green;

            // First label will be show Category name
            IDataLabel lbl = series.DataPoints[0].Label;
            lbl.DataLabelFormat.ShowCategoryName = true;

            lbl = series.DataPoints[1].Label;
            lbl.DataLabelFormat.ShowSeriesName = true;

            // Show value for third label
            lbl = series.DataPoints[2].Label;
            lbl.DataLabelFormat.ShowValue = true;
            lbl.DataLabelFormat.ShowSeriesName = true;
            lbl.DataLabelFormat.Separator = "/";
                        
            // Save presentation with chart
            pres.Save(dataDir + "AsposeChart_out.pptx", SaveFormat.Pptx);
            //ExEnd:NormalCharts
        }
    }
}