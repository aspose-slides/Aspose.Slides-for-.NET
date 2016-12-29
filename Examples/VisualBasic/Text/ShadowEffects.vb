Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Effects
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Text
    Public Class ShadowEffects
        Public Shared Sub Run()
            ' ExStart:ShadowEffects

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate a PPTX class
            Using pres As New Presentation()

                ' Get reference of the slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add an AutoShape of Rectangle type
                Dim ashp As IAutoShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50)


                ' Add TextFrame to the Rectangle
                ashp.AddTextFrame("Aspose TextBox")

                ' Disable shape fill in case we want to get shadow of text
                ashp.FillFormat.FillType = FillType.NoFill

                ' Add outer shadow and set all necessary parameters
                ashp.EffectFormat.EnableOuterShadowEffect()
                Dim shadow As IOuterShadow = ashp.EffectFormat.OuterShadowEffect
                shadow.BlurRadius = 4.0
                shadow.Direction = 45
                shadow.Distance = 3
                shadow.RectangleAlign = RectangleAlignment.TopLeft
                shadow.ShadowColor.PresetColor = PresetColor.Black

                'Write the presentation to disk
                pres.Save(dataDir & "pres_out.pptx", SaveFormat.Pptx)
            End Using
            ' ExEnd:ShadowEffects
        End Sub
    End Class
End Namespace