'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO
Imports System.Drawing.Imaging

Imports Aspose.Slides

Namespace CreateSlideThumbnail
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)

			'Getting the thumbnail image of the slide of a specified size
			Dim image As Image = slide.GetThumbnail(New Size(290, 230))

			'Saving the thumbnail image in jpeg format
			image.Save(dataDir & "thumbnail.jpg", ImageFormat.Jpeg)

			' Display Status.
			System.Console.WriteLine("Thumbnail created successfully.")
		End Sub
	End Class
End Namespace