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

Namespace ManagingFontFamilyOfText
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate PresentationEx Class
			Dim pres As New PresentationEx()


			'Get first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Add an AutoShape of Rectangle type
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 50, 50, 200, 50)

			'Access the added AutoShape
			Dim ashp As AutoShapeEx = CType(sld.Shapes(idx), AutoShapeEx)

			'Remove any fill style associated with the AutoShape
			ashp.FillFormat.FillType = FillTypeEx.NoFill

			'Access the TextFrame associated with the AutoShape
			Dim tf As TextFrameEx = ashp.TextFrame
			tf.Text = "Aspose TextBox"

			'Access the Portion associated with the TextFrame
			Dim port As PortionEx = tf.Paragraphs(0).Portions(0)

			'Set the Font for the Portion
			port.PortionFormat.LatinFont = New FontDataEx("Times New Roman")

			'Set Bold property of the Font 
			port.PortionFormat.FontBold = NullableBool.True

			'Set Italic property of the Font
			port.PortionFormat.FontItalic = NullableBool.True

			'Set Underline property of the Font 
			port.PortionFormat.FontUnderline = TextUnderlineTypeEx.Single

			'Set the Height of the Font
			port.PortionFormat.FontHeight = 25

			'Set the color of the Font
			port.PortionFormat.FillFormat.FillType = FillTypeEx.Solid
			port.PortionFormat.FillFormat.SolidFillColor.Color = Color.Blue

			'Write the presentation to disk
			pres.Write(dataDir & "output.pptx")


		End Sub
	End Class
End Namespace