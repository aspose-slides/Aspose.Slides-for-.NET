'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports Aspose.Slides
Imports Aspose.Slides.Export
Imports VisualBasic

Namespace ProgrammersGuide.Presentations
    Public Class ConvertSpecificSlideToPDF
        Public Shared Sub Run()

            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As Presentation = New Presentation(dataDir & "SelectedSlides.pptx")

                ' Setting array of slides positions
                Dim slides As Integer() = New Integer() {1, 3}

                ' Save the presentation to PDF
                presentation.Save(dataDir & "RequiredSelectedSlides.pdf", slides, SaveFormat.Pdf)
            End Using
        End Sub
    End Class
End Namespace