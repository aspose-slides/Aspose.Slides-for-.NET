Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.SlideShow
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class SimpleSlideTransitions
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate Presentation class that represents a presentation file
            Using pres As New Presentation(dataDir & "SimpleSlideTransitions.pptx")

                ' Apply circle type transition on slide 1
                pres.Slides(0).SlideShowTransition.Type = TransitionType.Circle

                ' Apply comb type transition on slide 2
                pres.Slides(1).SlideShowTransition.Type = TransitionType.Comb

                'Write the presentation to disk
                pres.Save(dataDir & "SampleTransition_out.pptx", SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace