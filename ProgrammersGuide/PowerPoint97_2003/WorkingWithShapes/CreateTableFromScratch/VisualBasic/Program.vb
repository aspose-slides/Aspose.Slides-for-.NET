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

Namespace CreateTableFromScratch
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)


			'Setting table parameters
			Dim xPosition As Integer = 880
			Dim yPosition As Integer = 1400
			Dim tableWidth As Integer = 4000
			Dim tableHeight As Integer = 500
			Dim columns As Integer = 4
			Dim rows As Integer = 4
			Dim borderWidth As Double = 2


			'Adding a new table to the slide using specified table parameters
			Dim table As Table = slide.Shapes.AddTable(xPosition, yPosition, tableWidth, tableHeight, columns, rows, borderWidth, Color.Blue)


			'Setting the alternative text for the table that can help in future to find
			'and identify this specific table
			table.AlternativeText = "myTable"


			'Merging two cells
			table.MergeCells(table.GetCell(0, 0), table.GetCell(1, 0))


			'Accessing the text frame of the first cell of the first row in the table
			Dim tf As TextFrame = table.GetCell(0, 0).TextFrame


			'If text frame is not null then add some text to the cell
			If tf IsNot Nothing Then
				tf.Paragraphs(0).Portions(0).Text = "Welcome to Aspose.Slides "
			End If


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")

		End Sub
	End Class
End Namespace