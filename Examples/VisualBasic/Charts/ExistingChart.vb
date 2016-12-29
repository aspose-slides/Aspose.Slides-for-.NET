Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class ExistingChart
        Public Shared Sub Run()
			'ExStart:ExistingChart	
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Instantiate Presentation class that represents PPTX file//Instantiate Presentation class that represents PPTX file
            Dim pres As New Presentation(dataDir & "ExistingChart.pptx")

            ' Access first slide
            Dim sld As ISlide = pres.Slides(0)

            ' Add chart with default data
            Dim chart As IChart = CType(sld.Shapes(0), IChart)

            ' Setting the index of chart data sheet
            Dim defaultWorksheetIndex As Integer = 0

            ' Getting the chart data worksheet
            Dim fact As IChartDataWorkbook = chart.ChartData.ChartDataWorkbook


            ' Changing chart Category Name
            fact.GetCell(defaultWorksheetIndex, 1, 0, "Modified Category 1")
            fact.GetCell(defaultWorksheetIndex, 2, 0, "Modified Category 2")


            'Take first chart series
            Dim series As IChartSeries = chart.ChartData.Series(0)

            ' Now updating series data
            fact.GetCell(defaultWorksheetIndex, 0, 1, "New_Series1") ' Modifying series name
            series.DataPoints(0).Value.Data = 90
            series.DataPoints(1).Value.Data = 123
            series.DataPoints(2).Value.Data = 44

            'Take Second chart series
            series = chart.ChartData.Series(1)

            ' Now updating series data
            fact.GetCell(defaultWorksheetIndex, 0, 2, "New_Series2") ' Modifying series name
            series.DataPoints(0).Value.Data = 23
            series.DataPoints(1).Value.Data = 67
            series.DataPoints(2).Value.Data = 99


            ' Now, Adding a new series
            chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 3, "Series 3"), chart.Type)

            'Take 3rd chart series
            series = chart.ChartData.Series(2)

            ' Now populating series data
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 1, 3, 20))
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 2, 3, 50))
            series.DataPoints.AddDataPointForBarSeries(fact.GetCell(defaultWorksheetIndex, 3, 3, 30))

            chart.Type = ChartType.ClusteredCylinder

            ' Save presentation with chart
            pres.Save(dataDir & "AsposeChartModified_out.pptx", SaveFormat.Pptx)

			'ExStart:ExistingChart
			
        End Sub
    End Class
End Namespace