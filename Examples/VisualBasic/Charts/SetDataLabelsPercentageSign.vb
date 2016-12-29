Imports System
Imports System.Drawing
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
' If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
' Install it and then add its reference to this project. For any issues, questions or suggestions 
' Please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class SetDataLabelsPercentageSign
        Public Shared Sub Run()
			'ExStart:SetDataLabelsPercentageSign
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation()

            ' Get reference of the slide
            Dim slide As ISlide = presentation.Slides(0)

            ' Add PercentsStackedColumn chart on a slide
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.PercentsStackedColumn, 20, 20, 500, 400)

            ' Set NumberFormatLinkedToSource to false
            chart.Axes.VerticalAxis.IsNumberFormatLinkedToSource = False
            chart.Axes.VerticalAxis.NumberFormat = "0.00%"

            chart.ChartData.Series.Clear()
            Dim defaultWorksheetIndex As Integer = 0

            ' Getting the chart data worksheet
            Dim workbook As IChartDataWorkbook = chart.ChartData.ChartDataWorkbook

            ' Add new series
            Dim series As IChartSeries = chart.ChartData.Series.Add(workbook.GetCell(defaultWorksheetIndex, 0, 1, "Reds"), chart.Type)
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 1, 1, 0.3))
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 2, 1, 0.5))
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 3, 1, 0.8))
            series.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 4, 1, 0.65))

            ' Setting the fill color of series
            series.Format.Fill.FillType = FillType.Solid
            series.Format.Fill.SolidFillColor.Color = Color.Red

            ' Setting LabelFormat properties
            series.Labels.DefaultDataLabelFormat.ShowValue = True
            series.Labels.DefaultDataLabelFormat.IsNumberFormatLinkedToSource = False
            series.Labels.DefaultDataLabelFormat.NumberFormat = "0.0%"
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FontHeight = 10
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid
            series.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.White
            series.Labels.DefaultDataLabelFormat.ShowValue = True

            ' Add new series
            Dim series2 As IChartSeries = chart.ChartData.Series.Add(workbook.GetCell(defaultWorksheetIndex, 0, 2, "Blues"), chart.Type)
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 1, 2, 0.7))
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 2, 2, 0.5))
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 3, 2, 0.2))
            series2.DataPoints.AddDataPointForBarSeries(workbook.GetCell(defaultWorksheetIndex, 4, 2, 0.35))

            ' Setting Fill type and color
            series2.Format.Fill.FillType = FillType.Solid
            series2.Format.Fill.SolidFillColor.Color = Color.Blue
            series2.Labels.DefaultDataLabelFormat.ShowValue = True
            series2.Labels.DefaultDataLabelFormat.IsNumberFormatLinkedToSource = False
            series2.Labels.DefaultDataLabelFormat.NumberFormat = "0.0%"
            series2.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FontHeight = 10
            series2.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.FillType = FillType.Solid
            series2.Labels.DefaultDataLabelFormat.TextFormat.PortionFormat.FillFormat.SolidFillColor.Color = Color.White

            ' Write presentation to disk
            presentation.Save(dataDir & Convert.ToString("SetDataLabelsPercentageSign_out.pptx"), SaveFormat.Pptx)
			'ExEnd:SetDataLabelsPercentageSign
        End Sub
    End Class
End Namespace