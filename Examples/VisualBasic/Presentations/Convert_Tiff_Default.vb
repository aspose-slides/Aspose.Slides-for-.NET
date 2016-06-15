Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Presentations
    Public Class Convert_Tiff_Default
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Presentations()

            'Instantiate a Presentation object that represents a presentation file
            Using presentation As New Presentation(dataDir & "Convert_Tiff_Default.pptx")

                'Saving the presentation to TIFF document
                presentation.Save(dataDir & "Tiff_out.tiff", Aspose.Slides.Export.SaveFormat.Tiff)
            End Using
        End Sub
    End Class
End Namespace