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
    public class SetDataLabelsPercentageSign
    {
        public static void Run()
        {
            //ExStart:SetDataLabelsPercentageSign
            // The path to the documents directory.
            string dataDir = RunExamples.GetDataDir_Charts();

            // Create an instance of Presentation class
            Presentation presentation = new Presentation();

            // Get reference of the slide
            ISlide slide = presentation.Slides[0];

            // Add PercentsStackedColumn chart on a slide
            IChart chart = slide.Shapes.AddChart(ChartType.PercentsStackedColumn, 20, 20, 500, 400);

            // Set NumberFormatLinkedToSource to false
            chart.Axes.VerticalAxis.IsNumberFormatLinkedToSource = false;
            chart.Axes.VerticalAxis.NumberFormat = "0.00%";

            chart.ChartData.Series.Clear();
            int defaultWorksheetIndex = 0;

            // Getting the chart data worksheet
            IChartDataWorkbook workbook = chart.ChartData.ChartDataWorkbook;

            // Add new series
            IChartSeries series = chart.ChartData.Series.Add(workbook.GetCell(defaultWorksheetIndex, 0, 1, "Reds"), chart.Type);
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 1, 1, 0.30));
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 2, 1, 0.50));
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 3, 1, 0.80));
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 4, 1, 0.65));

            // Setting the fill color of series
            series.Format.Fill.FillType = FillType.Solid;
            series.Format.Fill.SolidFillColor.Color = Color.Red;

            // Setting LabelFormat properties
            series.Labels.DefaultDataLabelFormat.ShowValue = true;
            series.Labels.DefaultDataLabelFormat.IsNumberFormatLinkedToSource = false;
            series.Labels.DefaultDataLabelFormat.NumberFormat = "0.0%";
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FontHeight = 10;
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.White;
            series.Labels.DefaultDataLabelFormat.ShowValue = true;

            // Add new series
            IChartSeries series2 = chart.ChartData.Series.Add(workbook.GetCell(defaultWorksheetIndex, 0, 2, "Blues"), chart.Type);
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 1, 2, 0.70));
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 2, 2, 0.50));
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 3, 2, 0.20));
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 4, 2, 0.35));

            // Setting Fill type and color
            series2.Format.Fill.FillType = FillType.Solid;
            series2.Format.Fill.SolidFillColor.Color = Color.Blue;
            series2.Labels.DefaultDataLabelFormat.ShowValue = true;
            series2.Labels.DefaultDataLabelFormat.IsNumberFormatLinkedToSource = false;
            series2.Labels.DefaultDataLabelFormat.NumberFormat = "0.0%";
            series2.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FontHeight = 10;
            series2.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid;
            series2.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.White;

            // Write presentation to disk
            presentation.Save(dataDir + "SetDataLabelsPercentageSign_out.pptx", SaveFormat.Pptx);
            //ExEnd:SetDataLabelsPercentageSign
        }
    }
}