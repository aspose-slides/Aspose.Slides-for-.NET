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

Namespace FindingShapeInSlide
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate a Presentation class that represents the presentation file
			Using p As New Presentation(dataDir & "SamplePres.pptx")

				Dim slide As ISlide = p.Slides(0)
				'alternative text of the shape to be found
				Dim shape As IShape = FindShape(slide, "Shape1")
				If shape IsNot Nothing Then
					Console.WriteLine("Shape Name: " & shape.Name)
				End If
			End Using
		End Sub

		'Method implementation to find a shape in a slide using its alternative text
		Public Shared Function FindShape(ByVal slide As ISlide, ByVal alttext As String) As IShape
			'Iterating through all shapes inside the slide
			For i As Integer = 0 To slide.Shapes.Count - 1


				'If the alternative text of the slide matches with the required one then
				'return the shape
				If slide.Shapes(i).AlternativeText.CompareTo(alttext) = 0 Then
					Return slide.Shapes(i)
				End If
			Next i
			Return Nothing
		End Function

	End Class
End Namespace

