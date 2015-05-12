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
Imports Aspose.Slides.SlideShow

Namespace ManagingBetterSlideTransitions
	Public Class Program
		Public Shared Sub Main(ByVal args() As String)
			' The path to the documents directory.
			Dim dataDir As String = Path.GetFullPath("../../../Data/")

			'Instantiate Presentation class that represents a presentation file
			Using pres As New Presentation(dataDir & "Aspose.pptx")

				'Apply circle type transition on slide 1
				pres.Slides(0).SlideShowTransition.Type = TransitionType.Circle


				'Set the transition time of 3 seconds
				pres.Slides(0).SlideShowTransition.AdvanceOnClick = True
				pres.Slides(0).SlideShowTransition.AdvanceAfterTime = 3000

				'Apply comb type transition on slide 2
				pres.Slides(1).SlideShowTransition.Type = TransitionType.Comb


				'Set the transition time of 5 seconds
				pres.Slides(1).SlideShowTransition.AdvanceOnClick = True
				pres.Slides(1).SlideShowTransition.AdvanceAfterTime = 5000

				'Apply zoom type transition on slide 3
				pres.Slides(2).SlideShowTransition.Type = TransitionType.Zoom


				'Set the transition time of 7 seconds
				pres.Slides(2).SlideShowTransition.AdvanceOnClick = True
				pres.Slides(2).SlideShowTransition.AdvanceAfterTime = 7000

				'Write the presentation to disk
				pres.Save(dataDir & "SampleTransition.pptx", SaveFormat.Pptx)

			End Using
		End Sub
	End Class
End Namespace