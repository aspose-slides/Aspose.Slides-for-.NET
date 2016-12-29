Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.SlideShow

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Transitions
    Public Class BetterSlideTransitions
        Public Shared Sub Run()
            ' ExStart:BetterSlideTransitions
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Transitions()

            ' Instantiate Presentation class that represents a presentation file
            Using pres As New Presentation(dataDir & Convert.ToString("BetterSlideTransitions.pptx"))

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

                ' Write the presentation to disk
                pres.Save(dataDir & Convert.ToString("SampleTransition_out.pptx"), SaveFormat.Pptx)
            End Using
            ' ExEnd:BetterSlideTransitions
        End Sub
    End Class
End Namespace