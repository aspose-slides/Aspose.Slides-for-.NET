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

Namespace CreatingTableFromScratchInSlideEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PresentationEx class that represents PPTX file
			Dim pres As New PresentationEx()

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Access first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Define columns with widths and rows with heights
			Dim dblCols() As Double = { 50, 50, 50 }
			Dim dblRows() As Double = { 50, 30, 30, 30, 30 }

			'Add table shape to slide
			Dim idx As Integer = sld.Shapes.AddTable(100, 50, dblCols, dblRows)
			Dim tbl As TableEx = CType(sld.Shapes(idx), TableEx)

			'Set border format for each cell
			For Each row As RowEx In tbl.Rows
				For Each cell As CellEx In row
					cell.BorderTop.FillFormat.FillType = FillTypeEx.Solid
					cell.BorderTop.FillFormat.SolidFillColor.Color = Color.Red
					cell.BorderTop.Width = 5

					cell.BorderBottom.FillFormat.FillType = FillTypeEx.Solid
					cell.BorderBottom.FillFormat.SolidFillColor.Color = Color.Red
					cell.BorderBottom.Width = 5

					cell.BorderLeft.FillFormat.FillType = FillTypeEx.Solid
					cell.BorderLeft.FillFormat.SolidFillColor.Color = Color.Red
					cell.BorderLeft.Width = 5

					cell.BorderRight.FillFormat.FillType = FillTypeEx.Solid
					cell.BorderRight.FillFormat.SolidFillColor.Color = Color.Red
					cell.BorderRight.Width = 5
				Next cell
			Next row

			'Merge cells 1 & 2 of row 1
			tbl.MergeCells(tbl(0, 0), tbl(1, 0), False)

			'Add text to the merged cell
			tbl(0, 0).TextFrame.Text = "Merged Cells"

			'Write PPTX to Disk
			pres.Write(dataDir & "table.pptx")

		End Sub
	End Class
End Namespace