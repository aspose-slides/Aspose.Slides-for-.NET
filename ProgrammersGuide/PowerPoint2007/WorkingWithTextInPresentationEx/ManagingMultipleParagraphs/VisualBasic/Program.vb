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
Imports System.Drawing

Namespace ManagingMultipleParagraphs
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate a PresentationEx class that represents a PPTX file
			Dim pres As New PresentationEx()


			'Accessing first slide
			Dim slide As SlideEx = pres.Slides(0)

			'Add an AutoShape of Rectangle type
			Dim idx As Integer = slide.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 50, 150, 300, 150)
			Dim ashp As AutoShapeEx = CType(slide.Shapes(idx), AutoShapeEx)

			'Access TextFrame of the AutoShape
			Dim tf As TextFrameEx = ashp.TextFrame

			'Create Paragraphs and Portions with different text formats
			Dim para0 As ParagraphEx = tf.Paragraphs(0)
			Dim port01 As New PortionEx()
			Dim port02 As New PortionEx()
			para0.Portions.Add(port01)
			para0.Portions.Add(port02)

			Dim para1 As New ParagraphEx()
			tf.Paragraphs.Add(para1)
			Dim port10 As New PortionEx()
			Dim port11 As New PortionEx()
			Dim port12 As New PortionEx()
			para1.Portions.Add(port10)
			para1.Portions.Add(port11)
			para1.Portions.Add(port12)

			Dim para2 As New ParagraphEx()
			tf.Paragraphs.Add(para2)
			Dim port20 As New PortionEx()
			Dim port21 As New PortionEx()
			Dim port22 As New PortionEx()
			para2.Portions.Add(port20)
			para2.Portions.Add(port21)
			para2.Portions.Add(port22)

			For i As Integer = 0 To 2
				For j As Integer = 0 To 2
					tf.Paragraphs(i).Portions(j).Text = "Portion0" & j.ToString()
					If j = 0 Then
						tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.FillType = FillTypeEx.Solid
						tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.SolidFillColor.Color = Color.Red
						tf.Paragraphs(i).Portions(j).PortionFormat.FontBold = NullableBool.True
						tf.Paragraphs(i).Portions(j).PortionFormat.FontHeight = 15
					ElseIf j = 1 Then
						tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.FillType = FillTypeEx.Solid
						tf.Paragraphs(i).Portions(j).PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue
						tf.Paragraphs(i).Portions(j).PortionFormat.FontItalic = NullableBool.True
						tf.Paragraphs(i).Portions(j).PortionFormat.FontHeight = 18
					End If
				Next j
			Next i

			'Write PPTX to Disk
			pres.Write(dataDir & "output.pptx")


		End Sub
	End Class
End Namespace