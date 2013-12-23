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

Namespace AccessExistingTable
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)


			'Setting table object to null
			Dim table As Table = Nothing


			'Iterating through all shapes unless the desired table is found
			For i As Integer = 0 To slide.Shapes.Count - 1
				If TypeOf slide.Shapes(i) Is Table Then
					table = CType(slide.Shapes(i), Table)


					If table.AlternativeText.Equals("myTable") Then
						System.Console.WriteLine("Table Found")
						Exit For
					End If
				End If
			Next i


			'Adding a new row in the table
			table.AddRow()


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")

		End Sub
	End Class
End Namespace