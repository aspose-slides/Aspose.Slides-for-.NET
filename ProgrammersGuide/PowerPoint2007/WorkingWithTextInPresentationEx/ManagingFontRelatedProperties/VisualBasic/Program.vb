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

Namespace ManagingFontRelatedProperties
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")


			'Accessing a slide using its slide position
			Dim slide As SlideEx = pres.Slides(0)

			'Accessing the first and second placeholder in the slide and typecasting it as AutoShape
			Dim tf1 As TextFrameEx = (CType(slide.Shapes(0), AutoShapeEx)).TextFrame
			Dim tf2 As TextFrameEx = (CType(slide.Shapes(1), AutoShapeEx)).TextFrame

			'Accessing the first Paragraph
			Dim para1 As ParagraphEx = tf1.Paragraphs(0)
			Dim para2 As ParagraphEx = tf2.Paragraphs(0)

			'Accessing the first portion
			Dim port1 As PortionEx = para1.Portions(0)
			Dim port2 As PortionEx = para2.Portions(0)

			'Define new fonts
			Dim fd1 As New FontDataEx("Elephant")
			Dim fd2 As New FontDataEx("Castellar")

			'Assign new fonts to portion
			port1.PortionFormat.LatinFont = fd1
			port2.PortionFormat.LatinFont = fd2

			'Set font to Bold
			port1.PortionFormat.FontBold = NullableBool.True
			port2.PortionFormat.FontBold = NullableBool.True

			'Set font to Italic
			port1.PortionFormat.FontItalic = NullableBool.True
			port2.PortionFormat.FontItalic = NullableBool.True

			'Set font color
			port1.PortionFormat.FillFormat.FillType = FillTypeEx.Solid
			port1.PortionFormat.FillFormat.SolidFillColor.Color = Color.Purple
			port2.PortionFormat.FillFormat.FillType = FillTypeEx.Solid
			port2.PortionFormat.FillFormat.SolidFillColor.Color = Color.Peru

			'Write the PPTX to disk
			pres.Write(dataDir & "output.pptx")


		End Sub
	End Class
End Namespace