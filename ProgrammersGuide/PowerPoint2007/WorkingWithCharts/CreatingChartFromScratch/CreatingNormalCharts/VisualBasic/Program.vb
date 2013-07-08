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

Namespace CreatingNormalCharts
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents PPTX file
			Dim pres As New PresentationEx()

			'Access first slide
			Dim sld As SlideEx = pres.Slides(0)

			' Add chart with default data
			Dim chart As ChartEx = sld.Shapes.AddChart(ChartTypeEx.ClusteredColumn, 0, 0, 500, 500)

			'Setting chart Title
			chart.ChartTitle.Text.Text = "Sample Title"
			chart.ChartTitle.Text.CenterText = True
			chart.ChartTitle.Height = 20
			chart.HasTitle = True

			'Set first series to Show Values
			chart.ChartData.Series(0).Labels.ShowValue = True

			'Setting the index of chart data sheet 
			Dim defaultWorksheetIndex As Integer = 0

			'Getting the chart data worksheet
			Dim fact As ChartDataCellFactory = chart.ChartData.ChartDataCellFactory

			'Delete default generated series and categories
			chart.ChartData.Series.Clear()
			chart.ChartData.Categories.Clear()
			Dim s As Integer = chart.ChartData.Series.Count
			s = chart.ChartData.Categories.Count

			'Adding new series
			chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 1, "Series 1"), chart.Type)
			chart.ChartData.Series.Add(fact.GetCell(defaultWorksheetIndex, 0, 2, "Series 2"), chart.Type)

			'Adding new categories
			chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 1, 0, "Caetegoty 1"))
			chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 2, 0, "Caetegoty 2"))
			chart.ChartData.Categories.Add(fact.GetCell(defaultWorksheetIndex, 3, 0, "Caetegoty 3"))

			'Take first chart series
			Dim series As ChartSeriesEx = chart.ChartData.Series(0)

			'Now populating series data
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, 20))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 2, 1, 50))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 3, 1, 30))

			'Setting fill color for series
			series.Format.Fill.FillType = FillTypeEx.Solid
			series.Format.Fill.SolidFillColor.Color = Color.Red


			'Take second chart series
			series = chart.ChartData.Series(1)

			'Now populating series data
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 1, 2, 30))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 2, 2, 10))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 3, 2, 60))

			'Setting fill color for series
			series.Format.Fill.FillType = FillTypeEx.Solid
			series.Format.Fill.SolidFillColor.Color = Color.Green


			'create custom labels for each of categories for new series

			'first label will be show Category name
			Dim lbl As New DataLabelEx(series)
			lbl.ShowCategoryName = True
			lbl.Id = 0
			series.Labels.Add(lbl)

			'Show series name for second label
			lbl = New DataLabelEx(series)
			lbl.ShowSeriesName = True
			lbl.Id = 1
			series.Labels.Add(lbl)

			'Show value for third label
			lbl = New DataLabelEx(series)
			lbl.ShowValue = True
			lbl.ShowSeriesName = True
			lbl.Separator = "/"
			lbl.Id = 2
			series.Labels.Add(lbl)

			'Show value and custom text
			lbl = New DataLabelEx(series)
			lbl.TextFrame.Text = "My text"
			lbl.Id = 3
			series.Labels.Add(lbl)

			'Save presentation with chart
			pres.Write(dataDir & "AsposeChart.pptx")


		End Sub
	End Class
End Namespace