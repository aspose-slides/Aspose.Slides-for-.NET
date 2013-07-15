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

Namespace FormattingChartEntities
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiating presentation
			Dim pres As New PresentationEx()

			'Accessing the first slide
			Dim slide As SlideEx = pres.Slides(0)

			'Adding the sample chart
			Dim chart As ChartEx = slide.Shapes.AddChart(ChartTypeEx.LineWithMarkers, 50, 50, 500, 400)

			'Setting Chart Titile
			chart.HasTitle = True
			Dim chartTitle As PortionEx = chart.ChartTitle.Text.Paragraphs(0).Portions(0)
			chartTitle.Text = "Sample Chart"
			chartTitle.PortionFormat.FillFormat.FillType = FillTypeEx.Solid
			chartTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray
			chartTitle.PortionFormat.FontHeight = 20
			chartTitle.PortionFormat.FontBold = NullableBool.True
			chartTitle.PortionFormat.FontItalic = NullableBool.True

			'Setting Major grid lines format for value axis
			chart.ValueAxis.MajorGridLines.FillFormat.FillType = FillTypeEx.Solid
			chart.ValueAxis.MajorGridLines.FillFormat.SolidFillColor.Color = Color.Blue
			chart.ValueAxis.MajorGridLines.Width = 5
			chart.ValueAxis.MajorGridLines.DashStyle = LineDashStyleEx.DashDot

			'Setting Minor grid lines format for value axis
			chart.ValueAxis.MinorGridLines.FillFormat.FillType = FillTypeEx.Solid
			chart.ValueAxis.MinorGridLines.FillFormat.SolidFillColor.Color = Color.Red
			chart.ValueAxis.MinorGridLines.Width = 3

			'Setting value axis number format
			chart.ValueAxis.SourceLinked = False
			chart.ValueAxis.DisplayUnit = DisplayUnitType.Thousands
			chart.ValueAxis.NumberFormat = "0.0%"

			'Setting chart maximum, minimum values
			chart.ValueAxis.IsAutomaticMajorUnit = False
			chart.ValueAxis.IsAutomaticMaxValue = False
			chart.ValueAxis.IsAutomaticMinorUnit = False
			chart.ValueAxis.IsAutomaticMinValue = False

			chart.ValueAxis.MaxValue = 15f
			chart.ValueAxis.MinValue = -2f
			chart.ValueAxis.MinorUnit = 0.5f
			chart.ValueAxis.MajorUnit = 2.0f

			'Setting Value Axis Text Properties
			Dim txtVal As TextFrameEx = chart.ValueAxis.TextProperties
			txtVal.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontBold = NullableBool.True
			txtVal.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontHeight = 16
			txtVal.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontItalic = NullableBool.True
			txtVal.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillTypeEx.Solid

			txtVal.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.DarkGreen
			txtVal.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.LatinFont = New FontDataEx("Times New Roman")

			'Setting value axis title
			chart.ValueAxis.HasTitle = True
			Dim valtitle As PortionEx = chart.ValueAxis.Title.Text.Paragraphs(0).Portions(0)
			valtitle.Text = "Primary Axis"
			valtitle.PortionFormat.FillFormat.FillType = FillTypeEx.Solid
			valtitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray
			valtitle.PortionFormat.FontHeight = 20
			valtitle.PortionFormat.FontBold = NullableBool.True
			valtitle.PortionFormat.FontItalic = NullableBool.True

			'Setting value axis line format
			chart.ValueAxis.Format.Line.Width = 10
			chart.ValueAxis.Format.Line.FillFormat.FillType = FillTypeEx.Solid
			chart.ValueAxis.Format.Line.FillFormat.SolidFillColor.Color = Color.Red

			'Setting Major grid lines format for Category axis
			chart.CategoryAxis.MajorGridLines.FillFormat.FillType = FillTypeEx.Solid
			chart.CategoryAxis.MajorGridLines.FillFormat.SolidFillColor.Color = Color.Green
			chart.CategoryAxis.MajorGridLines.Width = 5

			'Setting Minor grid lines format for Category axis
			chart.CategoryAxis.MinorGridLines.FillFormat.FillType = FillTypeEx.Solid
			chart.CategoryAxis.MinorGridLines.FillFormat.SolidFillColor.Color = Color.Yellow
			chart.CategoryAxis.MinorGridLines.Width = 3

			'Setting Category Axis Text Properties
			Dim txtCat As TextFrameEx = chart.CategoryAxis.TextProperties
			txtCat.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontBold = NullableBool.True
			txtCat.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontHeight = 16
			txtCat.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontItalic = NullableBool.True
			txtCat.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillTypeEx.Solid

			txtCat.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.Blue
			txtCat.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.LatinFont = New FontDataEx("Arial")

			'Setting Category Titile
			chart.CategoryAxis.HasTitle = True
			Dim catTitle As PortionEx = chart.CategoryAxis.Title.Text.Paragraphs(0).Portions(0)
			catTitle.Text = "Sample Category"
			catTitle.PortionFormat.FillFormat.FillType = FillTypeEx.Solid
			catTitle.PortionFormat.FillFormat.SolidFillColor.Color = Color.Gray
			catTitle.PortionFormat.FontHeight = 20
			catTitle.PortionFormat.FontBold = NullableBool.True
			catTitle.PortionFormat.FontItalic = NullableBool.True

			'Setting category axis lable position
			chart.CategoryAxis.TickLabelPosition = TickLabelPositionType.Low

			'Setting category axis lable rotation angle
			chart.CategoryAxis.RotationAngle = 45

			'Setting Legends Text Properties
			Dim txtleg As TextFrameEx = chart.Legend.TextProperties
			txtleg.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontBold = NullableBool.True
			txtleg.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontHeight = 16
			txtleg.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FontItalic = NullableBool.True
			txtleg.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FillFormat.FillType = FillTypeEx.Solid

			txtleg.Paragraphs(0).ParagraphFormat.DefaultPortionFormat.FillFormat.SolidFillColor.Color = Color.DarkRed

			'Set show chart legends without overlapping chart

			chart.Legend.Overlay = True

			'Setting secondary value axis
			chart.SecondValueAxis.IsVisible = True
			chart.SecondValueAxis.Format.Line.Style = LineStyleEx.ThickBetweenThin
			chart.SecondValueAxis.Format.Line.Width = 20

			'Setting secondary value axis Number format
			chart.SecondValueAxis.SourceLinked = False
			chart.SecondValueAxis.DisplayUnit = DisplayUnitType.Hundreds
			chart.SecondValueAxis.NumberFormat = "0.0%"

			'Setting chart maximum, minimum values
			chart.SecondValueAxis.IsAutomaticMajorUnit = False
			chart.SecondValueAxis.IsAutomaticMaxValue = False
			chart.SecondValueAxis.IsAutomaticMinorUnit = False
			chart.SecondValueAxis.IsAutomaticMinValue = False

			chart.SecondValueAxis.MaxValue = 20f
			chart.SecondValueAxis.MinValue = -5f
			chart.SecondValueAxis.MinorUnit = 0.5f
			chart.SecondValueAxis.MajorUnit = 2.0f

			'Ploting first series on secondary value axis
			chart.ChartData.Series(0).PlotOnSecondAxis = True

			'Setting chart back wall color
			chart.ChartFormat.Fill.FillType = FillTypeEx.Solid
			chart.ChartFormat.Fill.SolidFillColor.Color = Color.Orange

			'Setting Plot area color
			chart.PlotArea.Format.Fill.FillType = FillTypeEx.Solid
			chart.PlotArea.Format.Fill.SolidFillColor.Color = Color.LightCyan

			'Save Presentation
			pres.Write(dataDir & "ChartAxis.pptx")

		End Sub
	End Class
End Namespace