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

Namespace ManagingParagraphsAlignment
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")


			'Accessing first slide
			Dim slide As SlideEx = pres.Slides(0)

			'Accessing the first and second placeholder in the slide and typecasting it as AutoShape
			Dim tf1 As TextFrameEx = (CType(slide.Shapes(0), AutoShapeEx)).TextFrame
			Dim tf2 As TextFrameEx = (CType(slide.Shapes(1), AutoShapeEx)).TextFrame

			'Change the text in both placeholders
			tf1.Text = "Center Align by Aspose"
			tf2.Text = "Center Align by Aspose"

			'Getting the first paragraph of the placeholders
			Dim para1 As ParagraphEx = tf1.Paragraphs(0)
			Dim para2 As ParagraphEx = tf2.Paragraphs(0)

			'Aligning the text paragraph to center
			para1.ParagraphFormat.Alignment = TextAlignmentEx.Center
			para2.ParagraphFormat.Alignment = TextAlignmentEx.Center

			'Writing the presentation as a PPTX file
			pres.Write(dataDir & "output.pptx")


		End Sub
	End Class
End Namespace