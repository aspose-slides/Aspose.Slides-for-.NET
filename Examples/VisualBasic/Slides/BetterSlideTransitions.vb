Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export
Imports Aspose.Slides.SlideShow

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class BetterSlideTransitions
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate Presentation class that represents a presentation file
            Using pres As New Presentation(dataDir & "BetterSlideTransitions.pptx")

                ' Apply circle type transition on slide 1
                pres.Slides(0).SlideShowTransition.Type = TransitionType.Circle


                ' Set the transition time of 3 seconds
                pres.Slides(0).SlideShowTransition.AdvanceOnClick = True
                pres.Slides(0).SlideShowTransition.AdvanceAfterTime = 3000

                ' Apply comb type transition on slide 2
                pres.Slides(1).SlideShowTransition.Type = TransitionType.Comb


                ' Set the transition time of 5 seconds
                pres.Slides(1).SlideShowTransition.AdvanceOnClick = True
                pres.Slides(1).SlideShowTransition.AdvanceAfterTime = 5000

                ' Apply zoom type transition on slide 3
                pres.Slides(2).SlideShowTransition.Type = TransitionType.Zoom


                ' Set the transition time of 7 seconds
                pres.Slides(2).SlideShowTransition.AdvanceOnClick = True
                pres.Slides(2).SlideShowTransition.AdvanceAfterTime = 7000

                'Write the presentation to disk
                pres.Save(dataDir & "SampleTransition_out.pptx", SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace