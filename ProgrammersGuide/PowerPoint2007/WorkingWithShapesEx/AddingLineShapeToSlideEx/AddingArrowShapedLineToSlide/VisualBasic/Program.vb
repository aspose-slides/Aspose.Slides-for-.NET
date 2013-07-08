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

Namespace AddingArrowShapedLineToSlide
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents the PPTX file
			Dim pres As New PresentationEx()

			'Get the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Add an autoshape of type line
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Line, 50, 150, 300, 0)
			Dim shp As ShapeEx = sld.Shapes(idx)

			'Apply some formatting on the line
			shp.LineFormat.Style = LineStyleEx.ThickBetweenThin
			shp.LineFormat.Width = 10

			shp.LineFormat.DashStyle = LineDashStyleEx.DashDot

			shp.LineFormat.BeginArrowheadLength = LineArrowheadLengthEx.Short
			shp.LineFormat.BeginArrowheadStyle = LineArrowheadStyleEx.Oval

			shp.LineFormat.EndArrowheadLength = LineArrowheadLengthEx.Long
			shp.LineFormat.EndArrowheadStyle = LineArrowheadStyleEx.Triangle

			shp.LineFormat.FillFormat.FillType = FillTypeEx.Solid
			shp.LineFormat.FillFormat.SolidFillColor.Color = System.Drawing.Color.Maroon

			'Write the PPTX to Disk
			pres.Write(dataDir & "LineShape.pptx")

		End Sub
	End Class
End Namespace