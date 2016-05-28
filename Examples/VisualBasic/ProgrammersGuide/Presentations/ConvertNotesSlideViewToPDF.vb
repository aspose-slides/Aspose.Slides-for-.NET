'This project uses Automatic Package Restore feature of NuGet to resolve Aspose.Slides for .NET API reference 
'when the project is build. Please check https://docs.nuget.org/consume/nuget-faq for more information. 
'If you do not wish to use NuGet, you can manually download Aspose.Slides for .NET API from http://www.aspose.com/downloads, 
'install it and then add its reference to this project. For any issues, questions or suggestions 
'please feel free to contact us using http://www.aspose.com/community/forums/default.aspx

Imports System
Imports System.IO
Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace VisualBasic.Presentations
    Public Class ConvertNotesSlideViewToPDF
        Public Shared Sub Run()

            ' Loading a presentation
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using presentation As Presentation = New Presentation(dataDir + "NotesFile.pptx")

                ' Saving the presentation to PDF notes
                presentation.Save(dataDir + "Pdf_Notes.pdf", SaveFormat.PdfNotes)

            End Using


        End Sub
    End Class
End Namespace