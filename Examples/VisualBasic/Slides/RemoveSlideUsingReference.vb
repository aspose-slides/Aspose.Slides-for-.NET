Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class RemoveSlideUsingReference
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & "RemoveSlideUsingReference.pptx")

                ' Accessing a slide using its index in the slides collection
                Dim slide As ISlide = pres.Slides(0)


                ' Removing a slide using its reference
                pres.Slides.Remove(slide)


                'Writing the presentation file
                pres.Save(dataDir & "modified_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
        End Sub
    End Class
End Namespace