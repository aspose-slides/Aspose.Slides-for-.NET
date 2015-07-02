'////////////////////////////////////////////////////////////////////////
' Copyright 2001-2013 Aspose Pty Ltd. All Rights Reserved.
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
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            'Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                'Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                'Add autoshape of rectangle type
                Dim shp As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 150, 75)

                'Set the fill color of the rectangle shape
                shp.FillFormat.FillType = FillType.Solid
                shp.FillFormat.SolidFillColor.Color = Color.White

                'Apply some formatting on the line of the rectangle
                shp.LineFormat.Style = LineStyle.ThickThin
                shp.LineFormat.Width = 7
                shp.LineFormat.DashStyle = LineDashStyle.Dash

                'set the color of the line of rectangle
                shp.LineFormat.FillFormat.FillType = FillType.Solid
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Blue

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectShpLn.pptx", SaveFormat.Pptx)

            End Using
        End Sub
    End Class
End Namespace