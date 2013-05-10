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
Imports System.Drawing

Namespace AddingLineShapeToSlide
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)


			'Adding a line shape into the slide with its start and end points
			slide.Shapes.AddLine(New Point(1000, 1400), New Point(4700, 1400))

			'Adding a line shape into the slide with its start and end points
			Dim shape As Shape = slide.Shapes.AddLine(New Point(1000, 2000), New Point(4700, 2000))


			'Setting the fore color of the line to blue
			shape.LineFormat.ForeColor = Color.Blue


			'Setting the width of the line to 10
			shape.LineFormat.Width = 10


			'Setting the line style to thick between thin
			shape.LineFormat.Style = LineStyle.ThickBetweenThin


			'Setting the dash style of the line to dash
			shape.LineFormat.DashStyle = LineDashStyle.Dash


			'Setting the style of the starting point of the line to oval
			shape.LineFormat.BeginArrowheadStyle = LineArrowheadStyle.Oval


			'Setting the style of the ending point of the line to open
			shape.LineFormat.EndArrowheadStyle = LineArrowheadStyle.Open


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")
		End Sub
	End Class
End Namespace