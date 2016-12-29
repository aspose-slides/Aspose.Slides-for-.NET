using System.Drawing;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides;

/*
This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
install it and then add its reference to this project. For any issues, questions or suggestions 
please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
*/

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class PieChart
    {
        public static void Run()
        {
            //ExStart:PieChart
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Instantiate Presentation class that represents PPTX file
            Presentation presentation = new Presentation();

            // Access first slide
            ISlide slides = presentation.Slides[0];

            // Add chart with default data
            IChart chart = slides.Shapes.AddChart(ChartType.Pie, 100, 100, 400, 400);

            // Setting chart Title
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

            // Adding new categories
            chart.ChartData.Categories.Add(fact.GetCell(0, 1, 0, "First Qtr"));
            chart.ChartData.Categories.Add(fact.GetCell(0, 2, 0, "2nd Qtr"));
            chart.ChartData.Categories.Add(fact.GetCell(0, 3, 0, "3rd Qtr"));

            // Adding new series
            IChartSeries series = chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, "Series 1"), chart.Type);

            // Now populating series data
            series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 20));
            series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 50));
            series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 30));

            // Not working in new version
            // Adding new points and setting sector color
            // series.IsColorVaried = true;
            chart.ChartData.SeriesGroups[0].IsColorVaried = true;

            IChartDataPoint point = series.DataPoints[0];
            point.Format.Fill.FillType = FillType.Solid;
            point.Format.Fill.SolidFillColor.Color = Color.Cyan;
            // Setting Sector border
            point.Format.Line.FillFormat.FillType = FillType.Solid;
            point.Format.Line.FillFormat.SolidFillColor.Color = Color.Gray;
            point.Format.Line.Width = 3.0;
            point.Format.Line.Style = LineStyle.ThinThick;
            point.Format.Line.DashStyle = LineDashStyle.DashDot;

            IChartDataPoint point1 = series.DataPoints[1];
            point1.Format.Fill.FillType = FillType.Solid;
            point1.Format.Fill.SolidFillColor.Color = Color.Brown;

            // Setting Sector border
            point1.Format.Line.FillFormat.FillType = FillType.Solid;
            point1.Format.Line.FillFormat.SolidFillColor.Color = Color.Blue;
            point1.Format.Line.Width = 3.0;
            point1.Format.Line.Style = LineStyle.Single;
            point1.Format.Line.DashStyle = LineDashStyle.LargeDashDot;

            IChartDataPoint point2 = series.DataPoints[2];
            point2.Format.Fill.FillType = FillType.Solid;
            point2.Format.Fill.SolidFillColor.Color = Color.Coral;

            // Setting Sector border
            point2.Format.Line.FillFormat.FillType = FillType.Solid;
            point2.Format.Line.FillFormat.SolidFillColor.Color = Color.Red;
            point2.Format.Line.Width = 2.0;
            point2.Format.Line.Style = LineStyle.ThinThin;
            point2.Format.Line.DashStyle = LineDashStyle.LargeDashDotDot;

            // Create custom labels for each of categories for new series
            IDataLabel lbl1 = series.DataPoints[0].Label;

            // lbl.ShowCategoryName = true;
            lbl1.DataLabelFormat.ShowValue = true;

            IDataLabel lbl2 = series.DataPoints[1].Label;
            lbl2.DataLabelFormat.ShowValue = true;
            lbl2.DataLabelFormat.ShowLegendKey = true;
            lbl2.DataLabelFormat.ShowPercentage = true;

            IDataLabel lbl3 = series.DataPoints[2].Label;
            lbl3.DataLabelFormat.ShowSeriesName = true;
            lbl3.DataLabelFormat.ShowPercentage = true;

            // Showing Leader Lines for Chart
            series.Labels.DefaultDataLabelFormat.ShowLeaderLines = true;

            // Setting Rotation Angle for Pie Chart Sectors
            chart.ChartData.SeriesGroups[0].FirstSliceAngle = 180;

            // Save presentation with chart
            presentation.Save(dataDir + "PieChart_out.pptx", SaveFormat.Pptx);
            //ExEnd:PieChart
        }
    }
}