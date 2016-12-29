Imports System
Imports Aspose.Slides.Export
Imports Aspose.Slides.SlideShow

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Transitions
    Class SetTransitionEffects
        Public Shared Sub Run()
            ' ExStart:SetTransitionEffects
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Transitions()

            ' Create an instance of Presentation class
            Dim presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))

            ' Set effect
            presentation.Slides(0).SlideShowTransition.Type = TransitionType.Cut
            DirectCast(presentation.Slides(0).SlideShowTransition.Value, OptionalBlackTransition).FromBlack = True

            ' Write the presentation to disk
            presentation.Save(dataDir & Convert.ToString("SetTransitionEffects_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:SetTransitionEffects
        End Sub
    End Class
End Namespace