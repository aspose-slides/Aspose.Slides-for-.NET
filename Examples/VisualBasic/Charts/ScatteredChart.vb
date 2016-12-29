Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class ScatteredChart
        Public Shared Sub Run()
			'ExStart:ScatteredChart
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            Dim pres As New Presentation()

            Dim slide As ISlide = pres.Slides(0)

            ' Creating the default chart
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.ScatterWithSmoothLines, 0, 0, 400, 400)

            ' Getting the default chart data worksheet index
            Dim defaultWorksheetIndex As Integer = 0

            ' Getting the chart data worksheet
            Dim fact As IChartDataWorkbook = chart.ChartData.ChartDataWorkbook

            ' Delete demo series
            chart.ChartData.Series.Clear()

            ' Add new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, "Series 1"), chart.Type)
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 1, 3, "Series 2"), chart.Type)

            'Take first chart series
            Dim series As IChartSeries = chart.ChartData.Series(0)

            ' Add new point (1:3) there.
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 2, 1, 1), fact.GetCell(defaultWorksheetIndex, 2, 2, 3))

            ' Add new point (2:10)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 3, 1, 2), fact.GetCell(defaultWorksheetIndex, 3, 2, 10))

            ' Edit the type of series
            series.Type = ChartType.ScatterWithStraightLinesAndMarkers

            ' Changing the chart series marker
            series.Marker.Size = 10
            series.Marker.Symbol = MarkerStyleType.Star

            'Take second chart series
            series = chart.ChartData.Series(1)

            ' Add new point (5:2) there.
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 2, 3, 5), fact.GetCell(defaultWorksheetIndex, 2, 4, 2))

            ' Add new point (3:1)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 3, 3, 3), fact.GetCell(defaultWorksheetIndex, 3, 4, 1))

            ' Add new point (2:2)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 4, 3, 2), fact.GetCell(defaultWorksheetIndex, 4, 4, 2))

            ' Add new point (5:1)
            series.DataPoints.AddDataPointForScatterSeries(fact.GetCell(defaultWorksheetIndex, 5, 3, 5), fact.GetCell(defaultWorksheetIndex, 5, 4, 1))

            ' Changing the chart series marker
            series.Marker.Size = 10
            series.Marker.Symbol = MarkerStyleType.Circle

            pres.Save(dataDir & "AsposeChart_out.pptx", SaveFormat.Pptx)

			'ExEnd:ScatteredChart
        End Sub
    End Class
End Namespace