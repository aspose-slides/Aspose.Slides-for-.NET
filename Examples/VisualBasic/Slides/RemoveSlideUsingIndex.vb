Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class RemoveSlideUsingIndex
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & "RemoveSlideUsingIndex.pptx")

                'Removing a slide using its slide index
                pres.Slides.RemoveAt(0)


                'Writing the presentation file
                pres.Save(dataDir & "modified.pptx", Aspose.Slides.Export.SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace