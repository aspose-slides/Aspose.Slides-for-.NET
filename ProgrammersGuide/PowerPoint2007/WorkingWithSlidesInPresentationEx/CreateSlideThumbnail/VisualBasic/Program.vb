'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Pptx

Namespace CreateSlideThumbnail
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a PresentationEx class that represents the PPTX file
			Dim pres As New PresentationEx(dataDir & "demo.pptx")

			'Access the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Create a full scale image
			Dim bmp As Bitmap = sld.GetThumbnail(1F, 1F)

			'Save the image to disk in JPEG format
			bmp.Save(dataDir & "thumbnail.jpg", System.Drawing.Imaging.ImageFormat.Jpeg)

			' Display Status.
			System.Console.WriteLine("Thumbnail created successfully.")
		End Sub
	End Class
End Namespace