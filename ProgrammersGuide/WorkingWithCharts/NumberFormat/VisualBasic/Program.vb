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
Imports Aspose.Slides.Charts
Imports Aspose.Slides.Export

Namespace NumberFormat
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate the presentation//Instantiate the presentation
			Dim pres As New Presentation()

			'Access the first presentation slide
			Dim slide As ISlide = pres.Slides(0)

			'Adding a defautlt clustered column chart
			Dim chart As IChart = slide.Shapes.AddChart(ChartType.ClusteredColumn, 50, 50, 500, 400)

			'Accessing the chart series collection
			Dim series As IChartSeriesCollection = chart.ChartData.Series

			'Setting the preset number format
			'Traverse through every chart series
			For Each ser As ChartSeries In series
				'Traverse through every data cell in series
				For Each cell As IChartDataPoint In ser.DataPoints
					'Setting the number format
					cell.Value.AsCell.PresetNumberFormat = 10 '0.00%
				Next cell
			Next ser

			'Saving presentation
			pres.Save(dataDir & "PresetNumberFormat.pptx", SaveFormat.Pptx)

		End Sub
	End Class
End Namespace