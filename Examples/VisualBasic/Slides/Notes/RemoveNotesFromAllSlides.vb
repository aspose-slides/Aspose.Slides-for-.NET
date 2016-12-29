Imports System
Imports Aspose.Slides.Export

'
'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https:// Docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx
'

Namespace Aspose.Slides.Examples.VisualBasic.Slides.Notes
    Class RemoveNotesFromAllSlides
        Public Shared Sub Run()
            ' ExStart:RemoveNotesFromAllSlides
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Notes()

            ' Instantiate a Presentation object that represents a presentation file 
            Dim presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))

            ' Removing notes of all slides
            Dim mgr As INotesSlideManager = Nothing
            For i As Integer = 0 To presentation.Slides.Count - 1
                mgr = presentation.Slides(i).NotesSlideManager
                mgr.RemoveNotesSlide()
            Next

            presentation.Save(dataDir & Convert.ToString("RemoveNotesFromAllSlides_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:RemoveNotesFromAllSlides
        End Sub
    End Class
End Namespace