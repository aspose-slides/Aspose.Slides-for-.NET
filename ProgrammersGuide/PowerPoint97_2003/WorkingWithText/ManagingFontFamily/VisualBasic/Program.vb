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

Namespace ManagingFontFamily
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate Presentation class
			Dim pres As New Presentation()

			'Get first slide
			Dim sld As Slide = pres.GetSlideByPosition(1)

			'Add a Rectangle shape to the slide
			Dim rect As Aspose.Slides.Rectangle = sld.Shapes.AddRectangle(500, 500, 1500, 75)

			'Add a TextFrame to the Rectangle
			Dim tf As TextFrame = rect.AddTextFrame("Aspose Text Box")

			'Resize the Rectangle to fit text
			tf.FitShapeToText = True

			'Get the Portion object associated with the TextFrame
			Dim port As Portion = tf.Paragraphs(0).Portions(0)

			'Set the font of the portion
			pres.Fonts(port.FontIndex).FontName = "Times New Roman"

			'Set Bold property of the Font
			port.FontBold = True

			'Set Italic property of the Font
			port.FontItalic = True

			'Set Underline property of the Font
			port.FontUnderline = True

			'Set Height of the Font
			port.FontHeight = 25

			'Set the Color of the Font
			port.FontColor = Color.Blue

			'Write the Presentation to the disk
			pres.Write(dataDir & "output.ppt")

		End Sub
	End Class
End Namespace