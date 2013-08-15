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

Namespace ReplacingTextInPlaceholder
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents PPTX
			Dim pres As New PresentationEx(dataDir & "demo.pptx")


			'Access first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Iterate through shapes to find the placeholder
			For Each shp As ShapeEx In sld.Shapes
				If shp.Placeholder IsNot Nothing Then
					'Change the text of each placeholder
					CType(shp, AutoShapeEx).TextFrame.Text = "This is Placeholder"
				End If
			Next shp

			'Write the PPTX to Disk
			pres.Write(dataDir & "output.pptx")
		End Sub
	End Class
End Namespace