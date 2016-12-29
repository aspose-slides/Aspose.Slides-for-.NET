

Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports Aspose.Slides.Export
Imports System.Drawing

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class FillShapesPattern
        Public Shared Sub Run()
			'ExStart:FillShapesPattern
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate Prseetation class that represents the PPTX
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add autoshape of rectangle type
                Dim shp As IShape = sld.Shapes.AddAutoShape(ShapeType.Rectangle, 50, 150, 75, 150)

                ' Set the fill type to Pattern
                shp.FillFormat.FillType = FillType.Pattern

                ' Set the pattern style
                shp.FillFormat.PatternFormat.PatternStyle = PatternStyle.Trellis

                ' Set the pattern back and fore colors
                shp.FillFormat.PatternFormat.BackColor.Color = Color.LightGray
                shp.FillFormat.PatternFormat.ForeColor.Color = Color.Yellow

                'Write the PPTX file to disk
                pres.Save(dataDir & "RectShpPatt_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:FillShapesPattern
        End Sub
    End Class
End Namespace