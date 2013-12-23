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
Imports System

Namespace AddLinkedOLE
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			Dim pres As New Presentation()

			'Access the first slide
			Dim slide As Slide = pres.GetSlideByPosition(1)

			' Adding substitute image image. This is mandatory operation.
			Dim pic As New Picture(pres, dataDir & "logo.jpg")
			pres.Pictures.Add(pic)

			' Creating new linked Ole object
			Dim ole As OleObjectFrame = slide.Shapes.AddOleObjectFrame(500, 100, 500, 500, "Excel.Sheet.8", New Guid("{00020820-0000-0000-C000-000000000046}"), Nothing, dataDir & "book1.xls", Nothing)
			ole.PictureId = pic.PictureId

			' Replacing link and type of an object
			ole.SetObjectLink("Excel.Sheet.8", New Guid("{00020820-0000-0000-C000-000000000046}"), Nothing, dataDir & "book1.xls", Nothing)
			ole.PictureId = pic.PictureId

			' Replacing link path without changing of object's type
			ole.SetObjectLink(Nothing, dataDir & "book1.xls", Nothing)

			'Save presentation
			pres.Write(dataDir & "modified.ppt")

		End Sub
	End Class
End Namespace