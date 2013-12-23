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

Namespace ManagingParagraphIndent
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate Presentation Class
			Dim pres As New Presentation()

			'Get first slide
			Dim sld As Slide = pres.GetSlideByPosition(1)

			'Add a Rectangle Shape
			Dim rect As Aspose.Slides.Rectangle = sld.Shapes.AddRectangle(500, 500, 1500, 150)

			'Add TextFrame to the Rectangle
			Dim tf As TextFrame = rect.AddTextFrame("This is first line " & Constants.vbCr & " This is second line " & Constants.vbCr & " This is third line")

			'Set the text to fit the shape
			tf.FitShapeToText = True

			'Hide the lines of the Rectangle
			tf.LineFormat.ShowLines = False

			'Get first Paragraph in the TextFrame and set its Indent
			Dim para1 As Paragraph = tf.Paragraphs(0)
			para1.BulletOffset = 150

			'Get second Paragraph in the TextFrame and set its Indent
			Dim para2 As Paragraph = tf.Paragraphs(1)
			para2.BulletOffset = 250

			'Get third Paragraph in the TextFrame and set its Indent 
			Dim para3 As Paragraph = tf.Paragraphs(2)
			para3.BulletOffset = 350

			'Write the Presentation to disk
			pres.Write(dataDir & "output.ppt")

		End Sub
	End Class
End Namespace