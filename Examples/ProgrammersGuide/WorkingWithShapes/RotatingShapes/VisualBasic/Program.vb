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

Namespace RotatingShapes
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate PrseetationEx class that represents the PPTX
			Using pres As New Presentation()

				'Get the first slide
				Dim sld As ISlide = pres.Slides(0)

				'Add autoshape of rectangle type
				Dim shp As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150)

				'Rotate the shape to 90 degree
				shp.Rotation = 90

				'Write the PPTX file to disk
				pres.Save(dataDir & "RectShpRot.pptx", SaveFormat.Pptx)
			End Using
		End Sub
	End Class
End Namespace