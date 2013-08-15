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

Namespace CreatingTextBoxWithHyperlink
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate a PresentationEx class that represents a PPTX
			Dim pptxPresentation As New PresentationEx()

			'Get first slide
			Dim slide As SlideEx = pptxPresentation.Slides(0)

			'Add an AutoShape of Rectangle Type
			Dim index As Integer = slide.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 150, 150, 50)

			'Get a reference to the added shape 
			Dim pptxShape As Aspose.Slides.Pptx.ShapeEx = slide.Shapes(index)

			'Cast the shape to AutoShape
			Dim pptxAutoShape As AutoShapeEx = CType(pptxShape, AutoShapeEx)

			'Access TextFrame associated with the AutoShape
			Dim TextFrame As TextFrameEx = pptxAutoShape.TextFrame

			'Add some text to the frame
			TextFrame.Paragraphs(0).Portions(0).Text = "Aspose.Slides"

			'Set Hyperlink for the portion text
			Dim HLink As New HyperlinkEx("http://www.aspose.com")
			TextFrame.Paragraphs(0).Portions(0).HLinkClick = HLink

			'Save the PPTX Presentation
			pptxPresentation.Write(dataDir & "output.pptx")

		End Sub
	End Class
End Namespace