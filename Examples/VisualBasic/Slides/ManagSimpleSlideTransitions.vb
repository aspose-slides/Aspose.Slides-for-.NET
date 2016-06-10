Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.SlideShow
Imports Aspose.Slides

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace VisualBasic.Slides
    Class ManagSimpleSlideTransitions
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' ExStart:ManagSimpleSlideTransitions
            ' Instantiate Presentation class to load the source presentation file
            Using presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))
                ' Apply circle type transition on slide 1
                presentation.Slides(0).SlideShowTransition.Type = TransitionType.Circle

                ' Apply comb type transition on slide 2
                presentation.Slides(1).SlideShowTransition.Type = TransitionType.Comb

                ' ExEnd:ManagSimpleSlideTransitions
                ' Write the presentation to disk
                presentation.Save(dataDir & Convert.ToString("SampleTransition.pptx"), SaveFormat.Pptx)
            End Using
        End Sub
    End Class
End Namespace