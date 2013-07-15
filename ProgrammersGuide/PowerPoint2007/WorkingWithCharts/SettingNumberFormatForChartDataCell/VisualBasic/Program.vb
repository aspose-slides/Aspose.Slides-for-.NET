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

Namespace SettingNumberFormatForChartDataCell
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate the presentation
			Dim pres As New PresentationEx()

			'Access the first presentation slide
			Dim slide As SlideEx = pres.Slides(0)

			'Adding a defautlt clustered column chart
			Dim chart As ChartEx = slide.Shapes.AddChart(ChartTypeEx.ClusteredColumn, 50, 50, 500, 400)

			'Accessing the chart series collection
			Dim series As ChartSeriesExCollection = chart.ChartData.Series


			'Setting the preset number format

			'Traverse through every chart series
			For Each ser As ChartSeriesEx In series
				'Traverse through every data cell in series
				For Each cell As ChartDataCell In ser.Values
					'Setting the number format  
					cell.PresetNumberFormat = 10 '0.00%
				Next cell
			Next ser

			'Saving presentation
			pres.Write(dataDir & "PresetNumberFormat.pptx")


			'Now setting the custom number format

			'Traverse through every chart series
			For Each ser As ChartSeriesEx In series
				'Traverse through every data cell in series
				For Each cell As ChartDataCell In ser.Values
					'Setting the number format  
					cell.CustomNumberFormat = "0.00000"
				Next cell
			Next ser
			'Saving presentation
			pres.Write(dataDir & "CustomNumberFormat.pptx")

		End Sub
	End Class
End Namespace