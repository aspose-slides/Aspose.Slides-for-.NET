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
Imports Aspose.Slides.Pptx.Animation
Imports System.Drawing

Namespace ApplyingAnimationsOnShapesInsideSlideEx
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate PrseetationEx class that represents the PPTX
			Dim pres As New PresentationEx()
			Dim sld As SlideEx = pres.Slides(0)

			'Now create effect "PathFootball" for existing shape from scratch.
			Dim idx As Integer = sld.Shapes.AddAutoShape(ShapeTypeEx.Rectangle, 150, 150, 250, 25)
			Dim ashp As AutoShapeEx = CType(sld.Shapes(idx), AutoShapeEx)
			ashp.AddTextFrame("Animated TextBox")

			'Add PathFootBall animation effect
			Dim shape As ShapeEx = pres.Slides(0).Shapes(idx)
			pres.Slides(0).Timeline.MainSequence.AddEffect(shape, EffectTypeEx.PathFootball, EffectSubtypeEx.None, EffectTriggerTypeEx.AfterPrevious)

			'Create some kind of "button".
			Dim index As Integer = pres.Slides(0).Shapes.AddAutoShape(ShapeTypeEx.Bevel, 10, 10, 20, 20)
			Dim shapeTrigger As ShapeEx = pres.Slides(0).Shapes(index)

			'Create sequence of effects for this button.
			Dim seqInter As SequenceEx = pres.Slides(0).Timeline.InteractiveSequences.Add(shapeTrigger)

			'Create custom user path. Our object will be moved only after "button" click.
			Dim fxUserPath As EffectEx = seqInter.AddEffect(shape, EffectTypeEx.PathUser, EffectSubtypeEx.None, EffectTriggerTypeEx.OnClick)

			'Created path is empty so we should add commands for moving.
			Dim motionBhv As MotionEffectEx = (CType(fxUserPath.Behaviors(0), MotionEffectEx))
			Dim pts(0) As PointF
			pts(0) = New PointF(0.076f, 0.59f)
			motionBhv.Path.Add(MotionCommandPathTypeEx.LineTo, pts, MotionPathPointsTypeEx.Auto, True)
			pts(0) = New PointF(-0.076f, -0.59f)
			motionBhv.Path.Add(MotionCommandPathTypeEx.LineTo, pts, MotionPathPointsTypeEx.Auto, False)
			motionBhv.Path.Add(MotionCommandPathTypeEx.End, Nothing, MotionPathPointsTypeEx.Auto, False)

			'Write the presentation as PPTX to disk
			pres.Write(dataDir & "AnimExample.pptx")


		End Sub
	End Class
End Namespace