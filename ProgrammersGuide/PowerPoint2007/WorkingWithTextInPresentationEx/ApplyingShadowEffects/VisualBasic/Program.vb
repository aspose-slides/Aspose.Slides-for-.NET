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
Imports Aspose.Slides.Pptx
Imports Aspose.Slides.Effects

Namespace ApplyingShadowEffects
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			' Create directory if it is not already present.
			Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
			If (Not IsExists) Then
				System.IO.Directory.CreateDirectory(dataDir)
			End If

			'Instantiate a PPTX class
			Dim pres As New PresentationEx()


			'Get first slide
			Dim sld As SlideEx = pres.Slides(0)

			'Add an AutoShape of Rectangle type
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 75, 150, 50)
			Dim ashp As AutoShapeEx = CType(sld.Shapes(idx), AutoShapeEx)

			'Add TextFrame to the Rectangle
			ashp.AddTextFrame("Aspose TextBox")

			' Disable shape fill in case we want to get shadow of text.
			ashp.FillFormat.FillType = FillTypeEx.NoFill

			' Add outer shadow and set all necessary parameters
			Dim shadow As New OuterShadow()
			ashp.EffectFormat.OuterShadowEffect = shadow
			shadow.BlurRadius = 4.0
			shadow.Direction = 45
			shadow.Distance = 3
			shadow.RectangleAlign = RectangleAlignmentEx.TopLeft
			shadow.ShadowColor.PresetColor = PresetColorEx.Black

			'Write the presentation to disk
			pres.Write(dataDir & "output.pptx")


		End Sub
	End Class
End Namespace