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
    Class RemoveNotesAtSpecificSlide
        Public Shared Sub Run()
            ' ExStart:RemoveNotesAtSpecificSlide
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations_Notes()

            ' Instantiate a Presentation object that represents a presentation file 
            Dim presentation As New Presentation(dataDir & Convert.ToString("AccessSlides.pptx"))

            ' Removing notes of first slide
            Dim mgr As INotesSlideManager = presentation.Slides(0).NotesSlideManager
            mgr.RemoveNotesSlide()

            ' Save presentation to disk
            presentation.Save(dataDir & Convert.ToString("RemoveNotesAtSpecificSlide_out.pptx"), SaveFormat.Pptx)
            ' ExEnd:RemoveNotesAtSpecificSlide
        End Sub
    End Class
End Namespace