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

Namespace AccessingAnExistingTableInSlideEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents PPTX
			Dim pres As New PresentationEx(dataDir & "table.pptx")

			'Access the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Initialize null TableEx
			Dim tbl As TableEx = Nothing

			'Iterate through the shapes and set a reference to the table found
			For Each shp As ShapeEx In sld.Shapes
				If TypeOf shp Is TableEx Then
					tbl = CType(shp, TableEx)
				End If
			Next shp

			'Set the text of the first column of second row
			tbl(0, 1).TextFrame.Text = "New"

			'Write the PPTX to Disk
			pres.Write(dataDir & "table_updated.pptx")

		End Sub
	End Class
End Namespace