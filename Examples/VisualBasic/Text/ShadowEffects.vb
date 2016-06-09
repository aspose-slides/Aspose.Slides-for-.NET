'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2015 Aspose Pty Ltd. All Rights Reserved.
'
' This file is part of Aspose.Slides. The source code in this file
' is only intended as a supplement to the documentation, and is provided
' "as is", without warranty of any kind, either expressed or implied.
'////////////////////////////////////////////////////////////////////////

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Effects
Imports Aspose.Slides.Export

Namespace VisualBasic.Text
    Public Class ShadowEffects
        Public Shared Sub Run()
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Text()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate a PPTX class
            Using pres As New Presentation()

                'Get reference of the slide
                Dim sld As ISlide = pres.Slides(0)

                'Add an AutoShape of Rectangle type
                Dim ashp As IAutoShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 150, 75, 150, 50)


                'Add TextFrame to the Rectangle
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
                pres.Save(dataDir & "pres.pptx", SaveFormat.Pptx)
            End Using

        End Sub
    End Class
End Namespace