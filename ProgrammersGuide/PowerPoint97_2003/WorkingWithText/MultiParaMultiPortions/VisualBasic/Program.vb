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

Namespace MultiParaMultiPortions
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation()

			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)

			'Adding Rectangle Shape and setting its properties
			Dim rect As Aspose.Slides.Rectangle = slide.Shapes.AddRectangle(50, 50, 500, 200)
			Dim lf As LineFormat = rect.LineFormat
			lf.ShowLines = False
			Dim tf As TextFrame = rect.AddTextFrame("Default Text")
			tf.FitShapeToText = True

			'Creating Paragraphs and Portions
			Dim para0 As Paragraph = tf.Paragraphs(0)
			Dim port01 As New Portion()
			Dim port02 As New Portion()
			para0.Portions.Add(port01)
			para0.Portions.Add(port02)

			Dim para1 As New Paragraph()
			tf.Paragraphs.Add(para1)
			Dim port10 As New Portion()
			Dim port11 As New Portion()
			Dim port12 As New Portion()
			para1.Portions.Add(port10)
			para1.Portions.Add(port11)
			para1.Portions.Add(port12)

			Dim para2 As New Paragraph()
			tf.Paragraphs.Add(para2)
			Dim port20 As New Portion()
			Dim port21 As New Portion()
			Dim port22 As New Portion()
			para2.Portions.Add(port20)
			para2.Portions.Add(port21)
			para2.Portions.Add(port22)

			For i As Integer = 0 To 2
				For j As Integer = 0 To 2
					tf.Paragraphs(i).Portions(j).Text = "Portion0" & j.ToString()
					If j = 0 Then
						tf.Paragraphs(i).Portions(j).FontColor = Color.Red
					ElseIf j = 1 Then
						tf.Paragraphs(i).Portions(j).FontColor = Color.Blue
					End If
				Next j
			Next i
			pres.Write(dataDir & "modified.ppt")
		End Sub
	End Class
End Namespace