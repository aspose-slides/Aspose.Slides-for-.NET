Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.SmartArt

Namespace Aspose.Slides.Examples.VisualBasic.SmartArts
    Public Class CreateSmartArtShape
        Public Shared Sub Run()
            ' ExStart:CreateSmartArtShape
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_SmartArts()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If
            ' Instantiate the presentation
            Using pres As New Presentation()

                ' Access the presentation slide
                Dim slide As ISlide = pres.Slides(0)

                ' Add Smart Art Shape
                Dim smart As ISmartArt = slide.Shapes.AddSmartArt(0, 0, 400, 400, SmartArtLayoutType.BasicBlockList)

                ' Saving presentation
                pres.Save(dataDir & "SimpleSmartArt_out.pptx", Aspose.Slides.Export.SaveFormat.Pptx)
            End Using
            ' ExEnd:CreateSmartArtShape
        End Sub
    End Class
End Namespace