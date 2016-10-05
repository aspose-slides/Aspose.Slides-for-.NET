Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class ConvertWithNoteToTiff
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & "ConvertWithNoteToTiff.pptx")

                ' Saving the presentation to TIFF notes
                pres.Save(dataDir & "TestNotes_out.tiff", Aspose.Slides.Export.SaveFormat.TiffNotes)
            End Using
        End Sub
    End Class
End Namespace