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
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace VisualBasic.Shapes
    Public Class FormatLines
        Public Shared Sub Run()

            ' For complete examples and data files, please go to https://github.com/aspose-slides/Aspose.Slides-for-.NET            ' The path to the documents directory.

            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Instantiate Prseetation class that represents the PPTX
            Using presentation As New Presentation()

                ' Get the first slide
                Dim islide As ISlide = presentation.Slides(0)

                ' Add autoshape of rectangle type
                Dim ishape As IShape = islide.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 75)

                ' Set the fill color of the rectangle shape
                ishape.FillFormat.FillType = FillType.Solid
                ishape.FillFormat.SolidFillColor.Color = Color.White

                ' Apply some formatting on the line of the rectangle
                ishape.LineFormat.Style = LineStyle.ThickThin
                ishape.LineFormat.Width = 7
                ishape.LineFormat.DashStyle = LineDashStyle.Dash

                ' set the color of the line of rectangle
                ishape.LineFormat.FillFormat.FillType = FillType.Solid
                ishape.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue

                ' Write the PPTX file to disk
                presentation.Save(dataDir & "RectShpLn.pptx", SaveFormat.Pptx)

            End Using


        End Sub
    End Class
End Namespace