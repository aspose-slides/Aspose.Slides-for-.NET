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
    Public Class PieChart
        Public Shared Sub Run()
			'ExStart:PieChart
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Instantiate Presentation class that represents PPTX file
            Dim presentation As New Presentation()

            ' Access first slide
            Dim slides As ISlide = presentation.Slides(0)

            ' Add chart with default data
            Dim chart As IChart = slides.Shapes.AddChart(ChartType.Pie, 100, 100, 400, 400)

            ' Setting chart Title
            chart.ChartTitle.AddTextFrameForOverriding("Sample Title")
            chart.ChartTitle.TextFrameForOverriding.TextFrameFormat.CenterText = NullableBool.[True]
            chart.ChartTitle.Height = 20
            chart.HasTitle = True

            ' Set first series to Show Values
            chart.ChartData.Series(0).Labels.DefaultDataLabelFormat.ShowValue = True

            ' Setting the index of chart data sheet
            Dim defaultWorksheetIndex As Integer = 0

            ' Getting the chart data worksheet
            Dim fact As IChartDataWorkbook = chart.ChartData.ChartDataWorkbook

            ' Delete default generated series and categories

            chart.ChartData.Series.Clear()
            chart.ChartData.Categories.Clear()

            ' Adding new categories
            chart.ChartData.Categories.Add(fact.GetCell(0, 1, 0, "First Qtr"))
            chart.ChartData.Categories.Add(fact.GetCell(0, 2, 0, "2nd Qtr"))
            chart.ChartData.Categories.Add(fact.GetCell(0, 3, 0, "3rd Qtr"))

            ' Adding new series
            Dim series As IChartSeries = chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, "Series 1"), chart.Type)

            ' Now populating series data
            series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 1, 1, 20))
            series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 50))
            series.DataPoints.AddDataPointForPieSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 30))

            ' Not working in new version
            ' Adding new points and setting sector color
            ' series.IsColorVaried = true
            chart.ChartData.SeriesGroups(0).IsColorVaried = True

            Dim point As IChartDataPoint = series.DataPoints(0)
            point.Format.Fill.FillType = FillType.Solid
            point.Format.Fill.SolidFillColor.Color = Color.Cyan
            ' Setting Sector border
            point.Format.Line.FillFormat.FillType = FillType.Solid
            point.Format.Line.FillFormat.SolidFillColor.Color = Color.Gray
            point.Format.Line.Width = 3.0
            point.Format.Line.Style = LineStyle.ThinThick
            point.Format.Line.DashStyle = LineDashStyle.DashDot

            Dim point1 As IChartDataPoint = series.DataPoints(1)
            point1.Format.Fill.FillType = FillType.Solid
            point1.Format.Fill.SolidFillColor.Color = Color.Brown

            ' Setting Sector border
            point1.Format.Line.FillFormat.FillType = FillType.Solid
            point1.Format.Line.FillFormat.SolidFillColor.Color = Color.Blue
            point1.Format.Line.Width = 3.0
            point1.Format.Line.Style = LineStyle.[Single]
            point1.Format.Line.DashStyle = LineDashStyle.LargeDashDot

            Dim point2 As IChartDataPoint = series.DataPoints(2)
            point2.Format.Fill.FillType = FillType.Solid
            point2.Format.Fill.SolidFillColor.Color = Color.Coral

            ' Setting Sector border
            point2.Format.Line.FillFormat.FillType = FillType.Solid
            point2.Format.Line.FillFormat.SolidFillColor.Color = Color.Red
            point2.Format.Line.Width = 2.0
            point2.Format.Line.Style = LineStyle.ThinThin
            point2.Format.Line.DashStyle = LineDashStyle.LargeDashDotDot

            ' Create custom labels for each of categories for new series

            Dim lbl1 As IDataLabel = series.DataPoints(0).Label
            ' lbl.ShowCategoryName = true

            lbl1.DataLabelFormat.ShowValue = True

            Dim lbl2 As IDataLabel = series.DataPoints(1).Label
            lbl2.DataLabelFormat.ShowValue = True
            lbl2.DataLabelFormat.ShowLegendKey = True
            lbl2.DataLabelFormat.ShowPercentage = True

            Dim lbl3 As IDataLabel = series.DataPoints(2).Label
            lbl3.DataLabelFormat.ShowSeriesName = True
            lbl3.DataLabelFormat.ShowPercentage = True

            ' Showing Leader Lines for Chart
            series.Labels.DefaultDataLabelFormat.ShowLeaderLines = True

            ' Setting Rotation Angle for Pie Chart Sectors
            chart.ChartData.SeriesGroups(0).FirstSliceAngle = 180

            ' Save presentation with chart
            presentation.Save(dataDir & Convert.ToString("PieChart_out.pptx"), SaveFormat.Pptx)
			
			'ExEnd:PieChart
			
        End Sub
    End Class
End Namespace
