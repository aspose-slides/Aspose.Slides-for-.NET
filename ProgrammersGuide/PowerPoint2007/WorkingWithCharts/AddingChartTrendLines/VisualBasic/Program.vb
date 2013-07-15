'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Pptx
Imports Aspose.Slides.Pptx.Charts
Imports System.Drawing

Namespace AddingChartTrendLines
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Creating empty presentation
			Dim pres As New PresentationEx()

			'Creating a clustered column chart
			Dim chart As ChartEx = pres.Slides(0).Shapes.AddChart(ChartTypeEx.ClusteredColumn, 20, 20, 500, 400)

			'Adding Exponential trend line for chart series 1
			Dim tredLineExp As New TrendlineEx(chart.ChartData.Series(0))
			tredLineExp.TrendlineType = TrendlineTypeEx.Exponential
			tredLineExp.DisplayEquation = False
			tredLineExp.DisplayRSquaredValue = False
			chart.ChartData.Series(0).TrendLines.Add(tredLineExp)

			'Adding Linear trend line for chart series 1
			Dim tredLineLin As New TrendlineEx(chart.ChartData.Series(0))
			tredLineLin.TrendlineType = TrendlineTypeEx.Linear
			tredLineLin.Format.Line.FillFormat.FillType = FillTypeEx.Solid
			tredLineLin.Format.Line.FillFormat.SolidFillColor.Color = Color.Red
			chart.ChartData.Series(0).TrendLines.Add(tredLineLin)


			'Adding Logarithmic trend line for chart series 2
			Dim tredLineLog As New TrendlineEx(chart.ChartData.Series(1))
			tredLineLog.TrendlineType = TrendlineTypeEx.Logarithmic

			tredLineLog.TextFrame.Text = "New log trend line"
			chart.ChartData.Series(1).TrendLines.Add(tredLineLog)

			'Adding MovingAverage trend line for chart series 2
			Dim tredLineMovAvg As New TrendlineEx(chart.ChartData.Series(1))
			tredLineMovAvg.TrendlineType = TrendlineTypeEx.MovingAverage
			tredLineMovAvg.Period = 3
			tredLineMovAvg.TrendlineName = "New TrendLine Name"
			chart.ChartData.Series(1).TrendLines.Add(tredLineMovAvg)

			'Adding Polynomial trend line for chart series 3
			Dim tredLinePol As New TrendlineEx(chart.ChartData.Series(2))
			tredLinePol.TrendlineType = TrendlineTypeEx.Polynomial
			tredLinePol.Forward = 1
			tredLinePol.Order = 3
			chart.ChartData.Series(2).TrendLines.Add(tredLinePol)

			'Adding Power trend line for chart series 3
			Dim tredLinePower As New TrendlineEx(chart.ChartData.Series(2))
			tredLinePower.TrendlineType = TrendlineTypeEx.Power
			tredLinePower.Backward = 1
			chart.ChartData.Series(2).TrendLines.Add(tredLinePower)

			'Saving presentation
			pres.Write(dataDir & "TrendLines.pptx")


		End Sub
	End Class
End Namespace