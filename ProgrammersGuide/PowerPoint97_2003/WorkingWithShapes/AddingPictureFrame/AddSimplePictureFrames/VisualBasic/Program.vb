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

Namespace AddSimplePictureFrames
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)


			'Creating a picture object that will be used to fill the ellipse
			Dim pic As New Picture(pres, dataDir & "asp.jpg")


			'Adding the picture object to pictures collection of the presentation
			'After the picture object is added, the picture is given a uniqe picture Id
			Dim picId As Integer = pres.Pictures.Add(pic)


			'Calculating picture width and height
			Dim pictureWidth As Integer = pres.Pictures(picId - 1).Image.Width * 4
			Dim pictureHeight As Integer = pres.Pictures(picId - 1).Image.Height * 4


			'Calculating slide width and height
			Dim slideWidth As Integer = slide.Background.Width
			Dim slideHeight As Integer = slide.Background.Height


			'Calculating the width and height of picture frame
			Dim pictureFrameWidth As Integer = Convert.ToInt32(slideWidth \ 2 - pictureWidth \ 2)
			Dim pictureFrameHeight As Integer = Convert.ToInt32(slideHeight \ 2 - pictureHeight \ 2)


			'Adding picture frame to the slide
			slide.Shapes.AddPictureFrame(picId, pictureFrameWidth, pictureFrameHeight, pictureWidth, pictureHeight)


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")

		End Sub
	End Class
End Namespace