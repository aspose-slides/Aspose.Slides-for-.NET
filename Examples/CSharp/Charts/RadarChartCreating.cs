using System.Drawing;
using System.IO;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using Aspose.Slides;

namespace Aspose.Slides.Examples.CSharp.Charts
{
    public class RadarChartCreation
    {
        public static void Run()
        {
            string outPath = Path.Combine(RunExamples.OutPath, "RadarChart_Out.pptx");

            using (Presentation pres = new Presentation())
            {
                // Access first slide
                ISlide sld = pres.Slides[0];

                // Add Radar chart
                IChart ichart = sld.Shapes.AddChart(ChartType.Radar, 0, 0, 400, 400);

                // Setting the index of chart data sheet
                int defaultWorksheetIndex = 0;

                // Getting the chart data WorkSheet
                IChartDataWorkbook fact = ichart.ChartData.ChartDataWorkbook;

                // Set chart title
                ichart.ChartTitle.AddTextFrameForOverriding("Radar Chart");

                // Delete default generated series and categories
                ichart.ChartData.Categories.Clear();
                ichart.ChartData.Series.Clear();

                // Adding new categories
                ichart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"));
                ichart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, "Caetegoty 3"));
                ichart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 5"));
                ichart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 4, 0, "Caetegoty 7"));
                ichart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 5, 0, "Caetegoty 9"));
                ichart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 6, 0, "Caetegoty 11"));

                // Adding new series
                ichart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), ichart.Type);
                ichart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), ichart.Type);

                // Now populating series data
                IChartSeries series = ichart.ChartData.Series[0];
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 2.7));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 2.4));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 1.5));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 4, 1, 3.5));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 5, 1, 5));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 6, 1, 3.5));

                // Set series color
                series.Format.Line.FillFormat.FillType = FillType.Solid;
                series.Format.Line.FillFormat.SolidFillColor.Color = Color.DarkRed;

                // Now populating another series data
                series = ichart.ChartData.Series[1];
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 1, 2, 2.5));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 2, 2, 2.4));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 3, 2, 1.6));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 4, 2, 3.5));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 5, 2, 4));
                series.DataPoints.AddDataPointForRadarSeries(fact.GetCell(defaultWorksheetIndex, 6, 2, 3.6));

                // Set series color
                series.Format.Line.FillFormat.FillType = FillType.Solid;
                series.Format.Line.FillFormat.SolidFillColor.Color = Color.Orange;

                // Set legend position
                ichart.Legend.Position = LegendPositionType.Bottom;

                // Setting Category Axis Text Properties
                IChartPortionFormat txtCat = ichart.Axes.HorizontalAxis.TextFormat.PortionFormat;
                txtCat.FontBold = NullableBool.True;
                txtCat.FontHeight = 10;
                txtCat.FillFormat.FillType = FillType.Solid; ;
                txtCat.FillFormat.SolidFillColor.Color = Color.DimGray;
                txtCat.LatinFont = new FontData("Calibri");

                // Setting Legends Text Properties
                IChartPortionFormat txtleg = ichart.Legend.TextFormat.PortionFormat;
                txtleg.FontBold = NullableBool.True;
                txtleg.FontHeight = 10;
                txtleg.FillFormat.FillType = FillType.Solid; ;
                txtleg.FillFormat.SolidFillColor.Color = Color.DimGray;
                txtCat.LatinFont = new FontData("Calibri");

                // Setting Value Axis Text Properties
                IChartPortionFormat txtVal = ichart.Axes.VerticalAxis.TextFormat.PortionFormat;
                txtVal.FontBold = NullableBool.True;
                txtVal.FontHeight = 10;
                txtVal.FillFormat.FillType = FillType.Solid; ;
                txtVal.FillFormat.SolidFillColor.Color = Color.DimGray;
                txtVal.LatinFont = new FontData("Calibri");

                // Setting value axis number format
                ichart.Axes.VerticalAxis.IsNumberFormatLinkedToSource = false;
                ichart.Axes.VerticalAxis.NumberFormat = "\"$\"#,##0.00";

                // Setting chart major unit value
                ichart.Axes.VerticalAxis.IsAutomaticMajorUnit = false;
                ichart.Axes.VerticalAxis.MajorUnit = 1.25f;

                // Save generated presentation
                pres.Save(outPath, SaveFormat.Pptx);
            }
        }
    }
}