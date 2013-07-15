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

Namespace SettingPieChartSectorColors
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate PresentationEx class that represents PPTX file
			Dim pres As New PresentationEx()

			'Access first slide
			Dim sld As SlideEx = pres.Slides(0)

			' Add chart with default data
			Dim chart As Aspose.Slides.Pptx.ChartEx = sld.Shapes.AddChart(ChartTypeEx.Pie, 100, 100, 400, 400)

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
			Dim fact As ChartDataWorkbook = chart.ChartData.ChartDataWorkbook

			'Delete default generated series and categories

			chart.ChartData.Series.Clear()
			chart.ChartData.Categories.Clear()

			'Adding new categories
			chart.ChartData.Categories.Add(fact.GetCell(0, 1, 0, "First Qtr"))
			chart.ChartData.Categories.Add(fact.GetCell(0, 2, 0, "2nd Qtr"))
			chart.ChartData.Categories.Add(fact.GetCell(0, 3, 0, "3rd Qtr"))

			'Adding new series
			Dim Id As Integer = chart.ChartData.Series.Add(fact.GetCell(0, 0, 1, "Series 1"), chart.Type)

			'Accessing added series
			Dim series As ChartSeriesEx = chart.ChartData.Series(Id)

			'Now populating series data
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 1, 1, 20))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 2, 1, 50))
			series.Values.Add(fact.GetCell(defaultWorksheetIndex, 3, 1, 30))

			'Adding new points and setting sector color
			series.IsColorVaried = True
			Dim point As New ChartPointEx(series)
			point.Index = 0
			point.Format.Fill.FillType = FillTypeEx.Solid
			point.Format.Fill.SolidFillColor.Color = Color.Cyan
			'Setting Sector border
			point.Format.Line.FillFormat.FillType = FillTypeEx.Solid
			point.Format.Line.FillFormat.SolidFillColor.Color = Color.Gray
			point.Format.Line.Width = 3.0
			point.Format.Line.Style = LineStyleEx.ThinThick
			point.Format.Line.DashStyle = LineDashStyleEx.DashDot



			Dim point1 As New ChartPointEx(series)
			point1.Index = 1
			point1.Format.Fill.FillType = FillTypeEx.Solid
			point1.Format.Fill.SolidFillColor.Color = Color.Brown

			'Setting Sector border
			point1.Format.Line.FillFormat.FillType = FillTypeEx.Solid
			point1.Format.Line.FillFormat.SolidFillColor.Color = Color.Blue
			point1.Format.Line.Width = 3.0
			point1.Format.Line.Style = LineStyleEx.Single
			point1.Format.Line.DashStyle = LineDashStyleEx.LargeDashDot

			Dim point2 As New ChartPointEx(series)
			point2.Index = 2
			point2.Format.Fill.FillType = FillTypeEx.Solid
			point2.Format.Fill.SolidFillColor.Color = Color.Coral

			'Setting Sector border
			point2.Format.Line.FillFormat.FillType = FillTypeEx.Solid
			point2.Format.Line.FillFormat.SolidFillColor.Color = Color.Red
			point2.Format.Line.Width = 2.0
			point2.Format.Line.Style = LineStyleEx.ThinThin
			point2.Format.Line.DashStyle = LineDashStyleEx.LargeDashDotDot

			'Adding Series Points
			series.Points.Add(point)
			series.Points.Add(point1)
			series.Points.Add(point2)

			'Create custom labels for each of categories for new series

			Dim lbl As New DataLabelEx(series)
			' lbl.ShowCategoryName = true;
			lbl.ShowValue = True
			lbl.Id = 0
			series.Labels.Add(lbl)

			'Showing Leader Lines for Chart
			series.Labels.ShowLeaderLines = True

			'Setting Rotation Angle for Pie Chart Sectors
			chart.ChartData.Series(0).FirstSliceAngle = 180

			' Save presentation with chart
			pres.Write(dataDir & "AsposeChart.pptx")

		End Sub
	End Class
End Namespace