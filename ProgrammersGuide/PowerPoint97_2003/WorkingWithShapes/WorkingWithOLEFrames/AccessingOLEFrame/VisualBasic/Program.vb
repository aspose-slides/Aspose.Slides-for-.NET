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

Namespace AccessingOLEFrame
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)

			'Finding excel chart shape and obtaining its reference as OleObjectFrame
			Dim oof As OleObjectFrame = TryCast(slide.Shapes(1), OleObjectFrame)

			'Check, if OleObjectFrame is not null then read the object data of
			'OleObjectFrame as an array of bytes. Then write those bytes as an excel file
			If oof IsNot Nothing Then
				Dim fstr As New FileStream(dataDir & "book1.xls", FileMode.Create, FileAccess.Write)
				Dim buf() As Byte = oof.ObjectData
				fstr.Write(buf, 0, buf.Length)
				fstr.Flush()
				fstr.Close()
				System.Console.WriteLine("Excel OLE Object written as excel1.xls file")
			End If

		End Sub
	End Class
End Namespace