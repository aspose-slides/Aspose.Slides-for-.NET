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
Imports Aspose.Slides.Export

Namespace AddingOLEObjectFrame
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate Prseetation class that represents the PPTX
			Dim pres As New Presentation()

			'Access the first slide
			Dim sld As ISlide = pres.Slides(0)

			'Load an cel file to stream
			Dim fs As New FileStream(dataDir & "book1.xlsx", FileMode.Open, FileAccess.Read)
			Dim mstream As New MemoryStream()
			Dim buf(4095) As Byte

			Do
				Dim bytesRead As Integer = fs.Read(buf, 0, buf.Length)
				If bytesRead <= 0 Then
					Exit Do
				End If
				mstream.Write(buf, 0, bytesRead)
			Loop

			'Add an Ole Object Frame shape
			Dim oof As IOleObjectFrame = sld.Shapes.AddOleObjectFrame(0, 0, pres.SlideSize.Size.Width, pres.SlideSize.Size.Height, "Excel.Sheet.12", mstream.ToArray())

			'Write the PPTX to disk
			pres.Save(dataDir & "OleEmbed.pptx", SaveFormat.Pptx)


		End Sub
	End Class
End Namespace