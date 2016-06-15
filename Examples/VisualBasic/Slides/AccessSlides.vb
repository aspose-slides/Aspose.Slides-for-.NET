Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class AccessSlides
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate a Presentation object that represents a presentation file
            Using pres As New Presentation(dataDir & "AccessSlides.pptx")

                'Accessing a slide using its slide index
                Dim slide As ISlide = pres.Slides(0)

                System.Console.WriteLine("Slide Number: " & slide.SlideNumber)

            End Using
        End Sub
    End Class
End Namespace