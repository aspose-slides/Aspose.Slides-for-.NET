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

Namespace ExistingTable
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate Presentation class that represents PPTX//Instantiate Presentation class that represents PPTX
			Using pres As New Presentation(dataDir & "table.pptx")

				'Access the first slide
				Dim sld As ISlide = pres.Slides(0)

				'Initialize null TableEx
				Dim tbl As ITable = Nothing

				'Iterate through the shapes and set a reference to the table found
				For Each shp As IShape In sld.Shapes
					If TypeOf shp Is ITable Then
						tbl = CType(shp, ITable)
					End If
				Next shp

				'Set the text of the first column of second row
				tbl(0, 1).TextFrame.Text = "New"

				'Write the PPTX to Disk
				pres.Save(dataDir & "table1.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
			End Using

		End Sub
	End Class
End Namespace