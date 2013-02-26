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

Namespace AddPictureFrameToSlideEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PrseetationEx class that represents the PPTX
			Dim pres As New PresentationEx()

			'Get the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Instantiate the ImageEx class
			Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "asp.jpg"), System.Drawing.Image)
			Dim imgx As ImageEx = pres.Images.AddImage(img)

			'Add Picture Frame with height and width equivalent of Picture
			sld.Shapes.AddPictureFrame(ShapeTypeEx.Rectangle, 50, 150, imgx.Width, imgx.Height, imgx)

			'Write the PPTX file to disk
			pres.Write(dataDir & "RectPicFrame.pptx")

			' Display Status.
			System.Console.WriteLine("Picture Frame added successfully.")
		End Sub
	End Class
End Namespace