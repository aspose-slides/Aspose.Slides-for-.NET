Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class CloneWithInSamePresentation
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate Presentation class that represents a presentation file
            Using pres As New Presentation(dataDir & "CloneWithInSamePresentation.pptx")

                ' Clone the desired slide to the end of the collection of slides in the same presentation
                Dim slds As ISlideCollection = pres.Slides

                ' Clone the desired slide to the specified index in the same presentation
                slds.InsertClone(2, pres.Slides(1))

                'Write the modified presentation to disk
                pres.Save(dataDir & "Aspose_clone1_out.pptx", SaveFormat.Pptx)

            End Using

        End Sub
    End Class
End Namespace