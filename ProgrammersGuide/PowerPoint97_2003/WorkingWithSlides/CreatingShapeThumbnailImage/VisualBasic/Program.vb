'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO
Imports System.Drawing
Imports Aspose.Slides
Imports System.Drawing.Imaging

Namespace CreatingShapeThumbnailImage
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)


			'Iterate all shapes on a slide and create thumbnails
			Dim shapes As ShapeCollection = slide.Shapes
			For i As Integer = 0 To shapes.Count - 1
				Dim shape As Shape = shapes(i)


				'Getting the thumbnail image of the shape
				Dim img As Image = slide.GetThumbnail(New Object() { shape }, 1.0, 1.0, shape.ShapeRectangle)


				'Saving the thumbnail image in gif format
				img.Save(dataDir & "demo" & i & ".gif", ImageFormat.Gif)
			Next i
		End Sub
	End Class
End Namespace