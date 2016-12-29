Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Charts
    Public Class ChartTrendLines
        Public Shared Sub Run()
			'ExStart:ChartTrendLines
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Charts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Creating empty presentation
            Dim pres As New Presentation()

            ' Creating a clustered column chart
            Dim chart As IChart = pres.Slides(0).Shapes.AddChart(ChartType.ClusteredColumn, 20, 20, 500, 400)

            ' Adding ponential trend line for chart series 1
            Dim tredLinep As ITrendline = chart.ChartData.Series(0).TrendLines.Add(TrendlineType.Exponential)
            tredLinep.DisplayEquation = False
            tredLinep.DisplayRSquaredValue = False

            ' Adding Linear trend line for chart series 1
            Dim tredLineLin As ITrendline = chart.ChartData.Series(0).TrendLines.Add(TrendlineType.Linear)
            tredLineLin.TrendlineType = TrendlineType.Linear
            tredLineLin.Format.Line.FillFormat.FillType = FillType.Solid
            tredLineLin.Format.Line.FillFormat.SolidFillColor.Color = Color.Red


            ' Adding Logarithmic trend line for chart series 2
            Dim tredLineLog As ITrendline = chart.ChartData.Series(1).TrendLines.Add(TrendlineType.Logarithmic)
            tredLineLog.TrendlineType = TrendlineType.Logarithmic
            tredLineLog.AddTextFrameForOverriding("New log trend line")

            ' Adding MovingAverage trend line for chart series 2
            Dim tredLineMovAvg As ITrendline = chart.ChartData.Series(1).TrendLines.Add(TrendlineType.MovingAverage)
            tredLineMovAvg.TrendlineType = TrendlineType.MovingAverage
            tredLineMovAvg.Period = 3
            tredLineMovAvg.TrendlineName = "New TrendLine Name"

            ' Adding Polynomial trend line for chart series 3
            Dim tredLinePol As ITrendline = chart.ChartData.Series(2).TrendLines.Add(TrendlineType.Polynomial)
            tredLinePol.TrendlineType = TrendlineType.Polynomial
            tredLinePol.Forward = 1
            tredLinePol.Order = 3

            ' Adding Power trend line for chart series 3
            Dim tredLinePower As ITrendline = chart.ChartData.Series(1).TrendLines.Add(TrendlineType.Power)
            tredLinePower.TrendlineType = TrendlineType.Power
            tredLinePower.Backward = 1

            ' Saving presentation
            pres.Save(dataDir & "ChartTrendLines_out.pptx", SaveFormat.Pptx)

			'ExEnd:ChartTrendLines
			
        End Sub
    End Class
End Namespace