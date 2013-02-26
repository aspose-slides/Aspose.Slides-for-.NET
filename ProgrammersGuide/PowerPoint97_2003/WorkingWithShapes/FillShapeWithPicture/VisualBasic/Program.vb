'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////
Imports System.IO

Imports Aspose.Slides

Namespace FillShapeWithPicture
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")

			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(2)

			'Adding an ellipse shape into the slide by defining its X,Y position, width
			'and height
			Dim shape As Shape = slide.Shapes.AddEllipse(1400, 1200, 3000, 2000)

			'Setting the fill type of the ellipse to picture
			shape.FillFormat.Type = FillType.Picture

			'Creating a picture object that will be used to fill the ellipse
			Dim pic As New Picture(pres, dataDir & "demo.jpg")

			'Adding the picture object to pictures collection of the presentation
			'After the picture object is added, the picture is given a uniqe picture Id
			Dim picId As Integer = pres.Pictures.Add(pic)

			'Setting the picture Id of the shape fill to the Id of the picture object
			shape.FillFormat.PictureId = picId

			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")
		End Sub
	End Class
End Namespace