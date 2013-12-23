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

Namespace AddEmbeddedFrame
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)

			'Reading excel chart from the excel file and save as an array of bytes
			Dim fstro As New FileStream(dataDir & "book1.xls", FileMode.Open, FileAccess.Read)
			Dim b(fstro.Length - 1) As Byte
			fstro.Read(b, 0, CInt(Fix(fstro.Length)))

			'Inserting the excel chart as new OleObjectFrame to a slide
			Dim oof As OleObjectFrame = slide.Shapes.AddOleObjectFrame(0, 0, pres.SlideSize.Width, pres.SlideSize.Height, "Excel.Sheet.8", b)

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")

		End Sub
	End Class
End Namespace