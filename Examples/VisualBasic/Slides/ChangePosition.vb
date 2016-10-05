Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class ChangePosition
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            ' Instantiate Presentation class to load the source presentation file
            Using pres As New Presentation(dataDir & "ChangePosition.pptx")
                ' Get the slide whose position is to be changed
                Dim sld As ISlide = pres.Slides(0)

                ' Set the new position for the slide
                sld.SlideNumber = 2

                'Write the presentation to disk
                pres.Save(dataDir & "ChangePosition_out.pptx", SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace