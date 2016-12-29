Imports Microsoft.VisualBasic
Imports System.IO

Imports Aspose.Slides
Imports System.Drawing
Imports Aspose.Slides.Export

Namespace Aspose.Slides.Examples.VisualBasic.Shapes
    Public Class AddArrowShapedLineToSlide
        Public Shared Sub Run()
			'ExStart:AddArrowShapedLineToSlide
            ' The path to the documents directory.
            Dim dataDir As String = RunExamples.GetDataDir_Shapes()

            ' Create directory if it is not already present.
            Dim IsExists As Boolean = System.IO.Directory.Exists(dataDir)
            If (Not IsExists) Then
                System.IO.Directory.CreateDirectory(dataDir)
            End If

            ' Instantiate PresentationEx class that represents the PPTX file
            Using pres As New Presentation()

                ' Get the first slide
                Dim sld As ISlide = pres.Slides(0)

                ' Add an autoshape of type line
                Dim shp As IAutoShape = sld.Shapes.AddAutoShape(ShapeType.Line, 50, 150, 300, 0)

                ' Apply some formatting on the line
                shp.LineFormat.Style = LineStyle.ThickBetweenThin
                shp.LineFormat.Width = 10

                shp.LineFormat.DashStyle = LineDashStyle.DashDot

                shp.LineFormat.BeginArrowheadLength = LineArrowheadLength.Short
                shp.LineFormat.BeginArrowheadStyle = LineArrowheadStyle.Oval

                shp.LineFormat.EndArrowheadLength = LineArrowheadLength.Long
                shp.LineFormat.EndArrowheadStyle = LineArrowheadStyle.Triangle

                shp.LineFormat.FillFormat.FillType = FillType.Solid
                shp.LineFormat.FillFormat.SolidFillColor.Color = Color.Maroon

                'Write the PPTX to Disk
                pres.Save(dataDir & "LineShape2_out.pptx", SaveFormat.Pptx)
            End Using
			'ExEnd:AddArrowShapedLineToSlide
        End Sub
    End Class
End Namespace