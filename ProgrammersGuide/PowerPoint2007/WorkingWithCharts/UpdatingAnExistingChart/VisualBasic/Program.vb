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

Namespace UpdatingAnExistingChart
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents PPTX file
			Dim pres As New PresentationEx(dataDir & "AsposeChart.pptx")

			'Access first slide
			Dim sld As SlideEx = pres.Slides(0)

			' Add chart with default data
			Dim chart As ChartEx = CType(sld.Shapes(0), ChartEx)

			'Setting the index of chart data sheet 
			Dim defaultWorksheetIndex As Integer = 0

			'Getting the chart data worksheet
			Dim fact As ChartDataCellFactory = chart.ChartData.ChartDataCellFactory


			'Changing chart Category Name
			fact.GetCell(defaultWorksheetIndex, 1, 0, "Modified Category 1")
			fact.GetCell(defaultWorksheetIndex, 2, 0, "Modified Category 2")


			'Take first chart series
			Dim series As ChartSeriesEx = chart.ChartData.Series(0)

			'Now updating series data
			fact.GetCell(defaultWorksheetIndex, 0, 1, "New_Series1") 'modifying series name
			series.Values(0).Value = 90
			series.Values(1).Value = 123
			series.Values(2).Value = 44

			'Take Second chart series
			series = chart.ChartData.Series(1)

			'Now updating series data
			fact.GetCell(defaultWorksheetIndex, 0, 2, "New_Series2") 'modifying series name
			series.Values(0).Value = 23
			series.Values(1).Value = 67
			series.Values(2).Value = 99


			'Now, Adding a new series
			chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 3, "Series 3"), chart.Type)

			'Take 3rd chart series
			series = chart.ChartData.Series(2)

			'Now populating series data
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 1, 3, 20))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 2, 3, 50))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 3, 3, 30))


			chart.Type = ChartTypeEx.ClusteredCylinder

			' Save presentation with chart
			pres.Write(dataDir & "AsposeChartMoodified.pptx")

		End Sub
	End Class
End Namespace