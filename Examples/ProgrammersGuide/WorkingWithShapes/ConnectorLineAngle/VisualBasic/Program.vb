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
Imports System

Namespace ConnectorLineAngle
	Public Class Program
		Public Shared Sub Main(ByVal args As String())
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			Dim pres As Presentation = New Presentation(dataDir & "SamplePresLine.pptx")
			Dim slide As Slide = CType(pres.Slides(0), Slide)
			Dim shape As Shape
			Dim i As Integer = 0
			Do While i < slide.Shapes.Count
				Dim dir As Double = 0.0
				shape = CType(slide.Shapes(i), Shape)
				If TypeOf shape Is AutoShape Then
					Dim ashp As AutoShape = CType(shape, AutoShape)
					If ashp.ShapeType = ShapeType.Line Then
						dir = getDirection(ashp.Width, ashp.Height, ashp.Frame.FlipH, ashp.Frame.FlipV)
					End If
				ElseIf TypeOf shape Is Connector Then
					Dim ashp As Connector = CType(shape, Connector)
					dir = getDirection(ashp.Width, ashp.Height, ashp.Frame.FlipH, ashp.Frame.FlipV)
				End If

				Console.WriteLine(dir)
				i += 1
			Loop

		End Sub
		Public Shared Function getDirection(ByVal w As Single, ByVal h As Single, ByVal flipH As Boolean, ByVal flipV As Boolean) As Double
			Dim endLineX As Single
			If flipH Then
				endLineX = w * (-1)
			Else
				endLineX = w * (1)
			End If
			Dim endLineY As Single
			If flipV Then
				endLineY = h * (-1)
			Else
				endLineY = h * (1)
			End If
			Dim endYAxisX As Single = 0
			Dim endYAxisY As Single = h
			Dim angle As Double = (Math.Atan2(endYAxisY, endYAxisX) - Math.Atan2(endLineY, endLineX))
			If angle < 0 Then
			angle += 2 * Math.PI
			End If
			Return angle * 180.0 / Math.PI
		End Function
	End Class

End Namespace