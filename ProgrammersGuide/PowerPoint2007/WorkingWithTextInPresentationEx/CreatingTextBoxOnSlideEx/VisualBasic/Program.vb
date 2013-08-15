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

Namespace CreatingTextBoxOnSlideEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate PresentationEx
			Dim pres As New PresentationEx()

			'Get the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Add an AutoShape of Rectangle type
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 75, 150, 50)
			Dim ashp As AutoShapeEx = CType(sld.Shapes(idx), AutoShapeEx)

			'Add TextFrame to the Rectangle
			ashp.AddTextFrame(" ")

			'Accessing the text frame
			Dim txtFrame As TextFrameEx = ashp.TextFrame

			'Create the Paragraph object for text frame
			Dim para As ParagraphEx = txtFrame.Paragraphs(0)

			'Create Portion object for paragraph
			Dim portion As PortionEx = para.Portions(0)

			'Set Text
			portion.Text = "Aspose TextBox"

			'Write the presentation to disk
			pres.Write(dataDir & "output.pptx")

		End Sub
	End Class
End Namespace