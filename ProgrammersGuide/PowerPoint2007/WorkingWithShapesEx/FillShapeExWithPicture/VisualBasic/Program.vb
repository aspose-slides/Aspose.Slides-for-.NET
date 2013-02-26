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

Namespace FillShapeExWithPicture
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PrseetationEx class that represents the PPTX
			Dim pres As New PresentationEx()

			'Get the first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Add auto shape of rectangle type
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 50, 150, 75, 150)
			Dim shp As ShapeEx = sld.Shapes(idx)

			'Set the fill type to Picture
			shp.FillFormat.FillType = FillTypeEx.Picture

			'Set the picture fill mode
			shp.FillFormat.PictureFillFormat.PictureFillMode = PictureFillModeEx.Tile

			'Set the picture
			Dim img As System.Drawing.Image = CType(New Bitmap(dataDir & "asp.jpg"), System.Drawing.Image)
			Dim imgx As ImageEx = pres.Images.AddImage(img)
			shp.FillFormat.PictureFillFormat.Picture.Image = imgx

			'Write the PPTX file to disk
			pres.Write(dataDir & "RectShpPic.pptx")

			' Display Status.
			System.Console.WriteLine("Shape added and filled with image successfully.")
		End Sub
	End Class
End Namespace