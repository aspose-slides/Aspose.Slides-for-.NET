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

Namespace FindingShapeInSlideEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents the PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")

			'Get the first slide
			Dim slide As SlideEx = pres.Slides(0)

			'Calling FindShape method and passing the slide reference with the
			'alternative text of the shape to be found
			Dim shape As ShapeEx = FindShape(slide, "Slides")

			If shape IsNot Nothing Then
				System.Console.WriteLine("Shape Name: " & shape.Name)
				System.Console.WriteLine("Shape Height: " & shape.Height)
				System.Console.WriteLine("Shape Width: " & shape.Width)
			End If
		End Sub

		'Method implementation to find a shape in a slide using its alternative text
	   Public Shared Function FindShape(ByVal slide As SlideEx, ByVal alttext As String) As ShapeEx
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