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

Namespace CreatingSlidesSVGImage
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation class that represents the presentation file

			Using pres As New Presentation(dataDir & "Aspose.pptx")

				'Access the first slide
				Dim sld As ISlide = pres.Slides(0)

				'Create a memory stream object
				Dim SvgStream As New MemoryStream()

				'Generate SVG image of slide and save in memory stream
				sld.WriteAsSvg(SvgStream)
				SvgStream.Position = 0

				'Save memory stream to file
				Using fileStream As Stream = System.IO.File.OpenWrite(dataDir & "Aspose.svg")
					Dim buffer(8 * 1024 - 1) As Byte
					Dim len As Integer
					len = SvgStream.Read(buffer, 0, buffer.Length)
					Do While len > 0
						fileStream.Write(buffer, 0, len)
						len = SvgStream.Read(buffer, 0, buffer.Length)
					Loop

				End Using
				SvgStream.Close()
			End Using
		End Sub
	End Class
End Namespace