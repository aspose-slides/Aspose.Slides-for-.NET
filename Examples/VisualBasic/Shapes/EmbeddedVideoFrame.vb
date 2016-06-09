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
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class EmbeddedVideoFrame
        Public Shared Sub Run()

            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET            ' The path to the documents directory.

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Prseetation class that represents the PPTX//Instantiate Prseetation class that represents the PPTX
            Using presentation As New Presentation()

                ' Get the first slide
                Dim iSlide As ISlide = presentation.Slides(0)

                ' Add autoshape of ellipse type
                Dim ishape As IShape = iSlide.Shapes.AddAutoShape(ShapeType.Ellipse, 50, 150, 75, 150)

                ' Apply some gradiant formatting to ellipse shape
                ishape.FillFormat.FillType = FillType.Gradient
                ishape.FillFormat.GradientFormat.GradientShape = GradientShape.Linear

                ' Set the Gradient Direction
                ishape.FillFormat.GradientFormat.GradientDirection = GradientDirection.FromCorner2

                ' Add two Gradiant Stops
                ishape.FillFormat.GradientFormat.GradientStops.Add(CSng(1.0), PresetColor.Purple)
                ishape.FillFormat.GradientFormat.GradientStops.Add(CSng(0), PresetColor.Red)

                ' Write the PPTX file to disk
                presentation.Save(dataDir & "EllipseShpGrad.pptx", SaveFormat.Pptx)

            End Using

        End Sub
    End Class
End Namespace