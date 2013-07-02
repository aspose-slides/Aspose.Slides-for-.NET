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

Namespace ControllingAnimationOrder
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate a Presentation object that represents a PPT file
			Dim pres As New Presentation(dataDir & "demo.ppt")


			'Accessing a slide using its slide position
			Dim slide As Slide = pres.GetSlideByPosition(1)


			'Adding two shapes to the slide
			Dim shape1 As Shape = slide.Shapes.AddRectangle(1400, 1100, 3000, 2000)
			Dim shape2 As Shape = slide.Shapes.AddEllipse(2400, 1150, 1000, 1900)


			'Applying animation effects on both shapes
			shape1.AnimationSettings.EntryEffect = ShapeEntryEffect.Spiral
			shape2.AnimationSettings.EntryEffect = ShapeEntryEffect.BoxOut


			'Setting the animation order for both shapes. According to below order, shape2
			'will animate first and then the shape1
			shape1.AnimationSettings.AnimationOrder = 2
			shape2.AnimationSettings.AnimationOrder = 1


			'Writing the presentation as a PPT file
			pres.Write(dataDir & "modified.ppt")


		End Sub
	End Class
End Namespace