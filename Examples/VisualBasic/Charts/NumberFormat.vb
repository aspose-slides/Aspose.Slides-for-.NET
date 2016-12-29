Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class NumberFormat
        Public Shared Sub Run()
			'ExStart:NumberFormat
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate the presentation//Instantiate the presentation
            Dim pres As New Presentation()

            ' Access the first presentation slide
            Dim slide As ISlide = pres.Slides(0)

            ' Adding a defautlt clustered column chart
            Dim chart As IChart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400)

            ' Accessing the chart series collection
            Dim series As IChartSeriesCollection = chart.ChartData.Series

            ' Setting the preset number format
            'Traverse through every chart series
            For Each ser As ChartSeries In series
                'Traverse through every data cell in series
                For Each cell As IChartDataPoint In ser.DataPoints
                    ' Setting the number format
                    cell.Value.AsCell.PresetNumberFormat = 10 '0.00%
                Next cell
            Next ser

            ' Saving presentation
            pres.Save(dataDir & "PresetNumberFormat_out.pptx", SaveFormat.Pptx)

			'ExEnd:NumberFormat
        End Sub
    End Class
End Namespace