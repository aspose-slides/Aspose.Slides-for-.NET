Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Slides
    Public Class SetBackgroundToGradient
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Slides_Presentations()

            'Instantiate the Presentation class that represents the presentation file
            Using pres As New Presentation(dataDir & "SetBackgroundToGradient.pptx")

                'Apply Gradiant effect to the Background
                pres.Slides(0).Background.Type = BackgroundType.OwnBackground
                pres.Slides(0).Background.FillFormat.FillType = FillType.Gradient
                pres.Slides(0).Background.FillFormat.GradientFormat.TileFlip = TileFlip.FlipBoth

                'Write the presentation to disk
                pres.Save(dataDir & "ContentBG_Grad.pptx", SaveFormat.Pptx)

            End Using

        End Sub
    End Class
End Namespace